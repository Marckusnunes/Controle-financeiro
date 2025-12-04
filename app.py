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
    page_title="Super Conciliador v2.9 (Diagnóstico)",
    layout="wide",
    page_icon="💸",
    initial_sidebar_state="expanded"
)

# ==========================================
# 1. FUNÇÕES DE LIMPEZA
# ==========================================

def gerar_chave_padronizada(texto_conta):
    if not isinstance(texto_conta, str): return None
    texto_conta = texto_conta.strip()
    
    # Lógica Contábil (Códigos longos com ponto)
    if '.' in texto_conta and len(texto_conta) > 12:
        partes = texto_conta.split('.')
        maior_parte = ""
        for p in partes:
            limpo = re.sub(r'\D', '', p)
            if len(limpo) > len(maior_parte): maior_parte = limpo
        if len(maior_parte) > 4: texto_conta = maior_parte

    # Lógica Caixa (Ag/Op/Conta)
    elif '/' in texto_conta:
        texto_conta = texto_conta.split('/')[-1]
            
    parte_numerica = re.sub(r'\D', '', texto_conta)
    if not parte_numerica: return None
    
    # Retorna últimos 7 dígitos
    return parte_numerica[-7:].zfill(7)

def limpar_valor_monetario(valor_str):
    if not isinstance(valor_str, str): return 0.0
    valor_upper = valor_str.upper()
    # Verifica sinal negativo (D, DEB, -, ou parênteses de negativo)
    eh_negativo = 'D' in valor_upper or 'DEB' in valor_upper or '-' in valor_str or '(' in valor_str
    
    limpo = re.sub(r'[^\d,\.]', '', valor_str)
    
    try:
        if not limpo: return 0.0
        # Lógica Brasileira (ponto separa milhar, virgula separa decimal)
        if ',' in limpo and '.' in limpo:
             limpo = limpo.replace('.', '').replace(',', '.')
        elif ',' in limpo:
             limpo = limpo.replace(',', '.')
        
        valor_float = float(limpo)
        return -valor_float if eh_negativo else valor_float
    except ValueError:
        return 0.0

def formatar_moeda_br(valor):
    if pd.isna(valor): return "0,00"
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ==========================================
# 2. LEITURA DE PDF (Blindada)
# ==========================================

def extrair_pdf_melhorado(arquivo, tipo_extrato):
    try:
        doc = fitz.open(stream=arquivo.read(), filetype="pdf")
        texto_completo = ""
        for pag in doc:
            texto_completo += pag.get_text() + "\n"
        doc.close()
        
        linhas = texto_completo.split('\n')
        
        # --- 1. CONTA (Estratégia "Ninja") ---
        conta_encontrada = "N/A"
        
        # Padrões explícitos
        padroes_conta = [
            r"Conta:\s*(\d{4}\/\d{3,4}\/[\d\-]+)",        # Caixa
            r"Conta\s*Vinculada:\s*(\d{4}\/\d{3,4}\/[\d\-]+)", # Caixa Vinc
            r"Conta\s*Corrente\s*[:\s]*([\d\.\-\/]+)",    # BB
            r"Conta\s*[:\s]*([\d\.\-\/]+)",               # Genérico
            r"Agência.*?Conta.*?([\d\.\-]{5,})",          # Cabeçalho BB
            r"C\/C\s*[:\s]*([\d\.\-\/]+)"                 # Abreviação C/C
        ]
        
        for p in padroes_conta:
            match = re.search(p, texto_completo, re.IGNORECASE)
            if match:
                conta_raw = match.group(1).strip()
                # Valida se tem pelo menos 4 numeros
                if len(re.sub(r'\D', '', conta_raw)) > 4:
                    conta_encontrada = conta_raw
                    break
        
        # BUSCA DESESPERADA: Se não achou conta explícita, procura padrão 99999-9 no topo do arquivo
        if conta_encontrada == "N/A":
            # Pega só as primeiras 20 linhas para evitar pegar numeros aleatorios do meio
            cabecalho = "\n".join(linhas[:20])
            match_solto = re.search(r"(\d{4,6}-\d)", cabecalho)
            if match_solto:
                conta_encontrada = match_solto.group(1)

        # --- 2. VALORES ---
        saldo_final = 0.0
        rendimento_total = 0.0
        
        # Regex para moeda (1.000,00 ou 1000,00)
        regex_valor = r"(\d{1,3}(?:\.\d{3})*,\d{2}|\d{1,3}(?:,\d{3})*\.\d{2})"

        for i, linha in enumerate(linhas):
            linha_upper = linha.upper().strip()
            
            # --- SALDO ---
            gatilhos_saldo = [
                "SALDO FINAL", "SALDO TOTAL", "SALDO ATUAL", "SALDO EM", 
                "SALDO LÍQUIDO", "SALDO LIQUIDO", "SALDO BRUTO", 
                "VALOR LIQUIDO", "TOTAL DISPONIVEL", "POSICAO EM", "TOTAL EM COTAS",
                "S A L D O"
            ]
            ignorar = ["ANTERIOR", "BLOQUEADO", "PROVISORIO", "RENDIMENTO", "RENTABILIDADE"]
            
            if any(g in linha_upper for g in gatilhos_saldo) and not any(ign in linha_upper for ign in ignorar):
                # 1. Tenta na mesma linha
                match_val = re.search(regex_valor, linha_upper)
                if match_val:
                    # Verifica sinal negativo na mesma linha
                    sinal = "-" if " D" in linha_upper or "DEB" in linha_upper or "-" in linha_upper else ""
                    v = limpar_valor_monetario(f"{sinal}{match_val.group(0)}")
                    if v != 0: saldo_final = v
                
                # 2. Tenta na próxima linha (Crucial para BB)
                elif i + 1 < len(linhas):
                    linha_prox = linhas[i+1].upper().strip()
                    match_prox = re.search(regex_valor, linha_prox)
                    if match_prox:
                        # Assume positivo se estiver isolado na linha de baixo de "Saldo Atual"
                        v = limpar_valor_monetario(match_prox.group(0))
                        if v != 0: saldo_final = v

            # --- RENDIMENTOS (Investimentos) ---
            if tipo_extrato == 'INV':
                gatilhos_rend = ["RENDIMENTO BRUTO", "RENTABILIDADE", "RENDIMENTO NO MÊS", "RENDIMENTO LIQUIDO", "RENTAB."]
                
                # Evita pegar "Rentabilidade acumulada no ano" que é % e não R$
                if any(g in linha_upper for g in gatilhos_rend) and "ACUMULADO" not in linha_upper and "ANO" not in linha_upper:
                    
                    valor_capturado = 0.0
                    
                    # 1. Tenta na mesma linha
                    match_val = re.search(regex_valor, linha)
                    if match_val:
                        valor_capturado = limpar_valor_monetario(match_val.group(0))
                    
                    # 2. Tenta na próxima linha
                    elif i + 1 < len(linhas):
                        match_prox = re.search(regex_valor, linhas[i+1])
                        if match_prox:
                            valor_capturado = limpar_valor_monetario(match_prox.group(0))
                    
                    # Filtros de segurança
                    if valor_capturado != 0 and valor_capturado < 50000000: # Anti-ruído (evita pegar saldo como rendimento)
                         # Se achou "Líquido", ele tem prioridade sobre o Bruto?
                         # Vamos somar tudo. Geralmente o extrato tem um OU outro no resumo.
                         rendimento_total += valor_capturado

        # Fallback 1: Contas sem movimento
        if saldo_final == 0.0 and ("NAO HOUVE MOVIMENTO" in texto_completo.upper() or "SEM MOVIMENTO" in texto_completo.upper()):
             match_ant = re.search(r"(?:SALDO ANTERIOR|SALDO).*?(\d{1,3}(?:\.\d{3})*,\d{2})", texto_completo, re.IGNORECASE | re.DOTALL)
             if match_ant: saldo_final = limpar_valor_monetario(match_ant.group(1))

        # Fallback 2: Busca genérica por TOTAL no fim (Fundos de Investimento BB)
        if saldo_final == 0.0 and tipo_extrato == 'INV':
            match_last = re.findall(r"(?:TOTAL|SALDO|ATUAL|LÍQUIDO).*?(\d{1,3}(?:\.\d{3})*,\d{2})", texto_completo, re.IGNORECASE)
            if match_last: 
                # Pega o último valor monetário encontrado associado a palavras de saldo
                saldo_final = limpar_valor_monetario(match_last[-1])

        # Sanitiza o texto raw para não quebrar CSV (remove quebras de linha)
        texto_limpo = texto_completo[:300].replace('\n', ' ').replace(';', ',')

        return {
            "Conta": conta_encontrada,
            "Saldo": saldo_final,
            "Rendimento": rendimento_total,
            "Texto_Raw": texto_limpo
        }
    
    except Exception as e:
        return {"Conta": "Erro", "Saldo": 0.0, "Rendimento": 0.0, "Texto_Raw": str(e)}

# ==========================================
# 3. LEITURA CONTÁBIL
# ==========================================
def processar_contabil(arquivo, tipo='SALDO'):
    if arquivo is None: return pd.DataFrame()
    try:
        # Tenta header=1
        df = pd.read_csv(arquivo, encoding='latin-1', sep=';', header=1, dtype=str)
        col_chave = None
        possiveis_chaves = ['Domicílio bancário', 'Conta', 'Nº Conta', 'Descrição', 'Conta Contabil']
        for col in df.columns:
            for p in possiveis_chaves:
                if p.lower() in str(col).lower():
                    col_chave = col; break
        
        # Se falhar, tenta header=0
        if not col_chave:
            arquivo.seek(0)
            df = pd.read_csv(arquivo, encoding='latin-1', sep=';', header=0, dtype=str)
            for col in df.columns:
                for p in possiveis_chaves:
                    if p.lower() in str(col).lower():
                        col_chave = col; break
        
        if not col_chave: return pd.DataFrame()

        # Busca coluna de valor
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
                
                # Ajuste os códigos se necessário
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
        st.error("Erro na leitura do CSV de Saldos. Verifique o delimitador (;).")
        return pd.DataFrame(), pd.DataFrame()

    df_contabil = df_saldos
    if not df_rendim.empty:
        df_contabil = pd.merge(df_saldos, df_rendim, on='Chave Primaria', how='outer').fillna(0)
    else:
        df_contabil['Rendimento_Contabil'] = 0.0

    dados_banco = []
    log_leitura = []

    # Processa CC
    for f in lista_extratos_cc:
        res = extrair_pdf_melhorado(f, 'CC')
        chave = gerar_chave_padronizada(res['Conta'])
        # Loga TUDO, mesmo se chave for None, para você ver o erro
        log_leitura.append({'Arquivo': f.name, 'Conta Lida': res['Conta'], 'Chave Gerada': str(chave), 'Saldo': res['Saldo'], 'Rendimento': 0.0, 'Tipo': 'Conta Corrente', 'Raw': res['Texto_Raw']})
        if chave: 
            dados_banco.append({'Chave Primaria': chave, 'Saldo_Banco_CC': res['Saldo'], 'Saldo_Banco_Aplic': 0.0, 'Rendimento_Banco': 0.0})

    # Processa INV
    for f in lista_extratos_inv:
        res = extrair_pdf_melhorado(f, 'INV')
        chave = gerar_chave_padronizada(res['Conta'])
        # Loga TUDO
        log_leitura.append({'Arquivo': f.name, 'Conta Lida': res['Conta'], 'Chave Gerada': str(chave), 'Saldo': res['Saldo'], 'Rendimento': res['Rendimento'], 'Tipo': 'Investimento', 'Raw': res['Texto_Raw']})
        if chave: 
            dados_banco.append({'Chave Primaria': chave, 'Saldo_Banco_CC': 0.0, 'Saldo_Banco_Aplic': res['Saldo'], 'Rendimento_Banco': res['Rendimento']})

    df_log = pd.DataFrame(log_leitura)
    
    if dados_banco:
        df_banco = pd.DataFrame(dados_banco).groupby('Chave Primaria').sum().reset_index()
    else:
        df_banco = pd.DataFrame(columns=['Chave Primaria', 'Saldo_Banco_CC', 'Saldo_Banco_Aplic', 'Rendimento_Banco'])

    df_final = pd.merge(df_contabil, df_banco, on='Chave Primaria', how='outer').fillna(0)

    # Limpeza Descrição
    if 'Descrição' in df_final.columns:
        df_final['Descrição'] = df_final['Descrição'].fillna('CONTA SEM DESCRIÇÃO')
    else:
        df_final['Descrição'] = 'CONTA SEM DESCRIÇÃO'

    # Cálculo Diferenças
    df_final['Diferenca_Saldo_CC'] = df_final['Saldo_Contabil_CC'] - df_final['Saldo_Banco_CC']
    df_final['Diferenca_Saldo_Aplic'] = df_final['Saldo_Contabil_Aplic'] - df_final['Saldo_Banco_Aplic']
    df_final['Diferenca_Rendimento'] = df_final['Rendimento_Contabil'] - df_final['Rendimento_Banco']

    cols = ['Descrição', 'Chave Primaria', 
            'Saldo_Contabil_CC', 'Saldo_Banco_CC', 'Diferenca_Saldo_CC',
            'Saldo_Contabil_Aplic', 'Saldo_Banco_Aplic', 'Diferenca_Saldo_Aplic',
            'Rendimento_Contabil', 'Rendimento_Banco', 'Diferenca_Rendimento']
    colunas_finais = [c for c in cols if c in df_final.columns]
    
    return df_final[colunas_finais], df_log

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=True) 
        ws = writer.sheets['Sheet1']
        for i in range(1, 20):
             ws.column_dimensions[get_column_letter(i)].width = 18
    return output.getvalue()

# ==========================================
# 5. INTERFACE
# ==========================================
st.title("💸 Super Conciliador v2.9 (Diagnóstico)")
st.markdown("### ⚠️ Modo Diagnóstico Ativado: Verifique a Aba 'Log' para erros de leitura")

st.markdown("---")
col_cont1, col_cont2 = st.columns(2)
with col_cont1: f_saldos = st.file_uploader("📂 Saldos (CSV)", type='csv')
with col_cont2: f_rendim = st.file_uploader("📂 Rendimentos (CSV)", type='csv')

st.markdown("---")
col_bb, col_caixa = st.columns(2)
with col_bb:
    st.subheader("🔵 Banco do Brasil")
    f_bb_cc = st.file_uploader("BB - Conta Corrente", type='pdf', accept_multiple_files=True, key="bb_cc")
    f_bb_inv = st.file_uploader("BB - Aplicações", type='pdf', accept_multiple_files=True, key="bb_inv")
with col_caixa:
    st.subheader("🟠 Caixa")
    f_caixa_cc = st.file_uploader("Caixa - Conta Corrente", type='pdf', accept_multiple_files=True, key="cx_cc")
    f_caixa_inv = st.file_uploader("Caixa - Aplicações", type='pdf', accept_multiple_files=True, key="cx_inv")

if st.button("Executar Diagnóstico", type="primary"):
    tem_banco = (f_bb_cc or f_bb_inv or f_caixa_cc or f_caixa_inv)
    if f_saldos and tem_banco:
        with st.spinner("Processando..."):
            
            # Agrupa arquivos
            lista_final_cc = []
            if f_bb_cc: lista_final_cc.extend(f_bb_cc)
            if f_caixa_cc: lista_final_cc.extend(f_caixa_cc)
            
            lista_final_inv = []
            if f_bb_inv: lista_final_inv.extend(f_bb_inv)
            if f_caixa_inv: lista_final_inv.extend(f_caixa_inv)
            
            df_final, df_log = executar_processo(f_saldos, f_rendim, lista_final_cc, lista_final_inv)
            
            if not df_final.empty:
                st.success("Análise concluída!")
                
                # Setup Visual
                df_display = df_final.copy()
                mapa_colunas = {
                    'Descrição': ('Dados', 'Descrição'),
                    'Chave Primaria': ('Dados', 'Conta'),
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
                cols_existentes = [c for c in df_display.columns if c in mapa_colunas]
                df_display = df_display[cols_existentes]
                df_display.columns = pd.MultiIndex.from_tuples([mapa_colunas[c] for c in df_display.columns])
                
                numeric_cols = df_display.select_dtypes(include=['float', 'int']).columns
                df_formatado = df_display.copy()
                for col in numeric_cols:
                    df_formatado[col] = df_formatado[col].apply(formatar_moeda_br)

                # Abas
                tab1, tab2, tab3 = st.tabs(["📊 Resultado", "🚨 Divergências", "🕵️ Log (Crucial)"])
                
                with tab1:
                    st.dataframe(df_formatado, use_container_width=True)
                    st.download_button("Baixar Resultado", to_excel(df_display), "conciliacao_v29.xlsx")
                
                with tab2:
                    filtro = (df_final['Diferenca_Saldo_CC'].abs() > 0.01) | \
                             (df_final['Diferenca_Saldo_Aplic'].abs() > 0.01) | \
                             (df_final['Diferenca_Rendimento'].abs() > 0.01)
                    df_div = df_formatado[filtro]
                    if df_div.empty: st.info("Sem divergências!")
                    else: st.dataframe(df_div, use_container_width=True)
                
                with tab3:
                    st.markdown("""
                    **Instruções de Diagnóstico:**
                    1. Procure na lista abaixo o arquivo que deu erro (Ex: `7695-3.pdf`).
                    2. Veja a coluna **'Conta Lida'**. Se estiver 'N/A' ou 'Erro', o robô não achou o número da conta no cabeçalho.
                    3. Veja a coluna **'Saldo'**. Se estiver 0.00, ele não achou o texto 'SALDO ATUAL' ou similar.
                    """)
                    st.dataframe(df_log)
                    st.download_button("Baixar Log (Limpo)", df_log.to_csv(index=False, sep=';'), "log_leitura.csv")
            else:
                st.error("Erro crítico na geração da tabela.")
    else:
        st.warning("Faltam arquivos.")
