import streamlit as st
import pandas as pd
import re
import io
import os
import fitz  # PyMuPDF
from openpyxl.utils import get_column_letter

# ==========================================
# CONFIGURAÇÃO GERAL
# ==========================================
st.set_page_config(
    page_title="Sistema de Conciliação de Saldos Financeios",
    layout="wide",
    page_icon="✅",
    initial_sidebar_state="expanded"
)

# ==========================================
# 1. FUNÇÕES DE LIMPEZA E FORMATAÇÃO
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
    
    return parte_numerica[-7:].zfill(7)

def limpar_valor_monetario(valor_str):
    if not isinstance(valor_str, str): return 0.0
    valor_upper = valor_str.upper()
    eh_negativo = 'D' in valor_upper or 'DEB' in valor_upper or '-' in valor_str or '(' in valor_str
    
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
        
        linhas = texto_completo.split('\n')
        
        # --- CONTA ---
        conta_encontrada = "N/A"
        padroes_conta = [
            r"Conta:\s*(\d{4}\/\d{3,4}\/[\d\-]+)", 
            r"Conta\s*Vinculada:\s*(\d{4}\/\d{3,4}\/[\d\-]+)", 
            r"Conta\s*Corrente\s*[:\s]*([\d\.\-\/]+)",    
            r"Conta\s*[:\s]*([\d\.\-\/]+)",                
            r"Agência.*?Conta.*?([\d\.\-]{5,})",           
            r"C\/C\s*[:\s]*([\d\.\-\/]+)"                  
        ]
        
        for p in padroes_conta:
            match = re.search(p, texto_completo, re.IGNORECASE)
            if match:
                conta_raw = match.group(1).strip()
                if len(re.sub(r'\D', '', conta_raw)) > 4:
                    conta_encontrada = conta_raw
                    break
        
        if conta_encontrada == "N/A":
            cabecalho = "\n".join(linhas[:25]) 
            match_solto = re.search(r"(\d{4,6}-\d)", cabecalho)
            if match_solto: conta_encontrada = match_solto.group(1)

        # --- VALORES ---
        saldo_final = 0.0
        rendimento_total = 0.0
        regex_valor = r"(\d{1,3}(?:\.\d{3})*,\d{2}|\d{1,3}(?:,\d{3})*\.\d{2})"

        for i, linha in enumerate(linhas):
            linha_upper = linha.upper().strip()
            
            # SALDO
            gatilhos_saldo = ["SALDO FINAL", "SALDO TOTAL", "SALDO ATUAL", "SALDO EM", "SALDO LÍQUIDO", "SALDO BRUTO", "VALOR LIQUIDO", "TOTAL DISPONIVEL", "POSICAO EM", "TOTAL EM COTAS", "S A L D O"]
            ignorar = ["ANTERIOR", "BLOQUEADO", "PROVISORIO", "RENDIMENTO", "RENTABILIDADE"]
            
            if any(g in linha_upper for g in gatilhos_saldo) and not any(ign in linha_upper for ign in ignorar):
                match_val = re.search(regex_valor, linha_upper)
                if match_val:
                    sinal = "-" if " D" in linha_upper or "DEB" in linha_upper or "-" in linha_upper else ""
                    v = limpar_valor_monetario(f"{sinal}{match_val.group(0)}")
                    if v != 0: saldo_final = v
                elif i + 1 < len(linhas):
                    match_prox = re.search(regex_valor, linhas[i+1])
                    if match_prox:
                        v = limpar_valor_monetario(match_prox.group(0))
                        if v != 0: saldo_final = v

            # RENDIMENTOS
            if tipo_extrato == 'INV':
                gatilhos_rend = ["RENDIMENTO BRUTO", "RENTABILIDADE", "RENDIMENTO NO MÊS", "RENDIMENTO LIQUIDO", "RENTAB."]
                if any(g in linha_upper for g in gatilhos_rend) and "ACUMULADO" not in linha_upper and "ANO" not in linha_upper:
                    valor_capturado = 0.0
                    match_val = re.search(regex_valor, linha)
                    if match_val: valor_capturado = limpar_valor_monetario(match_val.group(0))
                    elif i + 1 < len(linhas):
                        match_prox = re.search(regex_valor, linhas[i+1])
                        if match_prox: valor_capturado = limpar_valor_monetario(match_prox.group(0))
                    
                    if valor_capturado != 0 and valor_capturado < 50000000:
                         rendimento_total += valor_capturado

        if saldo_final == 0.0 and ("NAO HOUVE MOVIMENTO" in texto_completo.upper() or "SEM MOVIMENTO" in texto_completo.upper()):
             match_ant = re.search(r"(?:SALDO ANTERIOR|SALDO).*?(\d{1,3}(?:\.\d{3})*,\d{2})", texto_completo, re.IGNORECASE | re.DOTALL)
             if match_ant: saldo_final = limpar_valor_monetario(match_ant.group(1))

        if saldo_final == 0.0 and tipo_extrato == 'INV':
            match_last = re.findall(r"(?:TOTAL|SALDO|ATUAL|LÍQUIDO).*?(\d{1,3}(?:\.\d{3})*,\d{2})", texto_completo, re.IGNORECASE)
            if match_last: saldo_final = limpar_valor_monetario(match_last[-1])

        texto_limpo = texto_completo[:300].replace('\n', ' ').replace(';', ',')
        return {"Conta": conta_encontrada, "Saldo": saldo_final, "Rendimento": rendimento_total, "Texto_Raw": texto_limpo}
    except Exception as e:
        return {"Conta": "Erro", "Saldo": 0.0, "Rendimento": 0.0, "Texto_Raw": str(e)}

def carregar_depara():
    """
    Carrega o arquivo DE-PARA e padroniza as chaves.
    Assume que o arquivo está na pasta 'depara' dentro do repositório.
    """
    # 1. Monta o caminho absoluto baseado na localização deste script
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    caminho_arquivo = os.path.join(diretorio_atual, "depara", "DEPARA_CONTAS BANCÁRIAS_CEF.xlsx")

    try:
        # Tenta carregar o arquivo usando o caminho construído
        df_depara = pd.read_excel(
            caminho_arquivo,
            sheet_name="2025_JUNHO (2)",
            dtype=str,
            engine='openpyxl'
        )
        
        # Padronização
        df_depara.columns = ['Conta Antiga', 'Conta Nova']
        
        # Verifica se a função gerar_chave_padronizada existe no escopo
        if 'gerar_chave_padronizada' in globals():
            df_depara['Chave Antiga'] = df_depara['Conta Antiga'].apply(gerar_chave_padronizada)
            df_depara['Chave Nova'] = df_depara['Conta Nova'].apply(gerar_chave_padronizada)
        else:
            st.error("Erro: A função 'gerar_chave_padronizada' não foi definida.")
            return pd.DataFrame()

        return df_depara

    except FileNotFoundError:
        st.warning(f"Aviso: Arquivo DE-PARA não encontrado em: '{caminho_arquivo}'. Verifique se a pasta e o arquivo existem no GitHub.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Erro ao ler o arquivo DE-PARA: {e}")
        return pd.DataFrame()

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
                if p.lower() in str(col).lower(): col_chave = col; break
        
        if not col_chave:
            arquivo.seek(0)
            df = pd.read_csv(arquivo, encoding='latin-1', sep=';', header=0, dtype=str)
            for col in df.columns:
                for p in possiveis_chaves:
                    if p.lower() in str(col).lower(): col_chave = col; break
        if not col_chave: return pd.DataFrame()

        col_valor = None
        possiveis_valores = ['Saldo Final', 'Saldo Atual', 'Movimento', 'Valor']
        for col in df.columns:
            if 'anterior' in str(col).lower(): continue
            for p in possiveis_valores:
                if p.lower() in str(col).lower(): col_valor = col; break
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
# 4. CONSOLIDAÇÃO E DE-PARA
# ==========================================
def executar_processo(file_saldos, file_rendim, lista_extratos_cc, lista_extratos_inv):
    # 1. Carrega CSVs
    df_saldos = processar_contabil(file_saldos, 'SALDO')
    df_rendim = processar_contabil(file_rendim, 'RENDIMENTO')
    
    if df_saldos.empty:
        st.error("Erro na leitura do CSV de Saldos.")
        return pd.DataFrame(), pd.DataFrame()

    # 2. LÓGICA DE-PARA (Aqui é aplicada a substituição)
    df_depara = carregar_depara()
    
    if not df_depara.empty:
        # Cria dicionário {Antiga: Nova}
        dicionario_depara = dict(zip(df_depara['Chave Antiga'], df_depara['Chave Nova']))
        
        # Aplica no Saldos
        df_saldos['Chave Primaria'] = df_saldos['Chave Primaria'].replace(dicionario_depara)
        
        # Aplica no Rendimentos (se houver)
        if not df_rendim.empty:
            df_rendim['Chave Primaria'] = df_rendim['Chave Primaria'].replace(dicionario_depara)
            
        st.toast(f"✅ De-Para aplicado: {len(df_depara)} contas convertidas.")

    # 3. Merge Contábil
    df_contabil = df_saldos
    if not df_rendim.empty:
        df_contabil = pd.merge(df_saldos, df_rendim, on='Chave Primaria', how='outer').fillna(0)
    else:
        df_contabil['Rendimento_Contabil'] = 0.0

    # 4. Leitura dos PDFs (Bancos)
    dados_banco = []
    log_leitura = []

    for f in lista_extratos_cc:
        res = extrair_pdf_melhorado(f, 'CC')
        chave = gerar_chave_padronizada(res['Conta'])
        log_leitura.append({'Arquivo': f.name, 'Conta Lida': res['Conta'], 'Chave Gerada': str(chave), 'Saldo': res['Saldo'], 'Rendimento': 0.0, 'Tipo': 'CC', 'Raw': res['Texto_Raw']})
        if chave: dados_banco.append({'Chave Primaria': chave, 'Saldo_Banco_CC': res['Saldo'], 'Saldo_Banco_Aplic': 0.0, 'Rendimento_Banco': 0.0})

    for f in lista_extratos_inv:
        res = extrair_pdf_melhorado(f, 'INV')
        chave = gerar_chave_padronizada(res['Conta'])
        log_leitura.append({'Arquivo': f.name, 'Conta Lida': res['Conta'], 'Chave Gerada': str(chave), 'Saldo': res['Saldo'], 'Rendimento': res['Rendimento'], 'Tipo': 'INV', 'Raw': res['Texto_Raw']})
        if chave: dados_banco.append({'Chave Primaria': chave, 'Saldo_Banco_CC': 0.0, 'Saldo_Banco_Aplic': res['Saldo'], 'Rendimento_Banco': res['Rendimento']})

    df_log = pd.DataFrame(log_leitura)
    if dados_banco:
        df_banco = pd.DataFrame(dados_banco).groupby('Chave Primaria').sum().reset_index()
    else:
        df_banco = pd.DataFrame(columns=['Chave Primaria', 'Saldo_Banco_CC', 'Saldo_Banco_Aplic', 'Rendimento_Banco'])

    # 5. Consolidação Final
    df_final = pd.merge(df_contabil, df_banco, on='Chave Primaria', how='outer').fillna(0)

    if 'Descrição' in df_final.columns: df_final['Descrição'] = df_final['Descrição'].fillna('CONTA SEM DESCRIÇÃO')
    else: df_final['Descrição'] = 'CONTA SEM DESCRIÇÃO'

    df_final['Diferenca_Saldo_CC'] = df_final['Saldo_Contabil_CC'] - df_final['Saldo_Banco_CC']
    df_final['Diferenca_Saldo_Aplic'] = df_final['Saldo_Contabil_Aplic'] - df_final['Saldo_Banco_Aplic']
    df_final['Diferenca_Rendimento'] = df_final['Rendimento_Contabil'] - df_final['Rendimento_Banco']

    cols = ['Descrição', 'Chave Primaria', 'Saldo_Contabil_CC', 'Saldo_Banco_CC', 'Diferenca_Saldo_CC',
            'Saldo_Contabil_Aplic', 'Saldo_Banco_Aplic', 'Diferenca_Saldo_Aplic',
            'Rendimento_Contabil', 'Rendimento_Banco', 'Diferenca_Rendimento']
    colunas_finais = [c for c in cols if c in df_final.columns]
    return df_final[colunas_finais], df_log

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=True) 
        ws = writer.sheets['Sheet1']
        for i in range(1, 20): ws.column_dimensions[get_column_letter(i)].width = 18
    return output.getvalue()

# ==========================================
# 5. INTERFACE (PAINEL DE STATUS)
# ==========================================
st.title("Sistema de Conciliação de Saldos Financeiros")

# PAINEL DE STATUS
st.markdown("### Status dos Arquivos")
col_cont1, col_cont2 = st.columns(2)
with col_cont1: f_saldos = st.file_uploader("📂 1. Saldos Contábeis (Relatório Flexvision 0113083)", type='csv')
with col_cont2: f_rendim = st.file_uploader("📂 2. Rendimentos (Relatório Flexvision 014387)", type='csv')

col_bb, col_caixa = st.columns(2)
with col_bb:
    f_bb_cc = st.file_uploader("🔵 Extrato BB - Conta Corrente", type='pdf', accept_multiple_files=True)
    f_bb_inv = st.file_uploader("🔵 Extrato BB - Aplicações", type='pdf', accept_multiple_files=True)
with col_caixa:
    f_caixa_cc = st.file_uploader("🟠 Extrato Caixa Econômica - Conta Corrente", type='pdf', accept_multiple_files=True)
    f_caixa_inv = st.file_uploader("🟠 Extrato Caixa Econômica - Aplicações", type='pdf', accept_multiple_files=True)

# LÓGICA DO STATUS
tem_csv = f_saldos is not None
tem_banco = (f_bb_cc or f_bb_inv or f_caixa_cc or f_caixa_inv)

c1, c2 = st.columns(2)
with c1:
    if tem_csv: st.success("✅ CSV de Saldos Carregado")
    else: st.error("❌ Faltando: Saldos (CSV)")
with c2:
    if tem_banco: st.success("✅ Pelo menos um PDF carregado")
    else: st.error("❌ Faltando: Extratos Bancários")

st.markdown("---")

if st.button("Executar Conciliação", type="primary"):
    if tem_csv and tem_banco:
        with st.spinner("Processando..."):
            lista_final_cc = []
            if f_bb_cc: lista_final_cc.extend(f_bb_cc)
            if f_caixa_cc: lista_final_cc.extend(f_caixa_cc)
            
            lista_final_inv = []
            if f_bb_inv: lista_final_inv.extend(f_bb_inv)
            if f_caixa_inv: lista_final_inv.extend(f_caixa_inv)
            
            df_final, df_log = executar_processo(f_saldos, f_rendim, lista_final_cc, lista_final_inv)
            
            if not df_final.empty:
                st.balloons()
                st.success("Sucesso!")
                
                df_display = df_final.copy()
                mapa_colunas = {
                    'Descrição': ('Dados', 'Descrição'), 'Chave Primaria': ('Dados', 'Conta'),
                    'Saldo_Contabil_CC': ('CONTA CORRENTE', 'Contabilidade'), 'Saldo_Banco_CC': ('CONTA CORRENTE', 'Extrato Banco'), 'Diferenca_Saldo_CC': ('CONTA CORRENTE', 'Diferença'),
                    'Saldo_Contabil_Aplic': ('APLICAÇÃO', 'Contabilidade'), 'Saldo_Banco_Aplic': ('APLICAÇÃO', 'Extrato Banco'), 'Diferenca_Saldo_Aplic': ('APLICAÇÃO', 'Diferença'),
                    'Rendimento_Contabil': ('RENDIMENTO', 'Contabilidade'), 'Rendimento_Banco': ('RENDIMENTO', 'Extrato Banco'), 'Diferenca_Rendimento': ('RENDIMENTO', 'Diferença')
                }
                cols_existentes = [c for c in df_display.columns if c in mapa_colunas]
                df_display = df_display[cols_existentes]
                df_display.columns = pd.MultiIndex.from_tuples([mapa_colunas[c] for c in df_display.columns])
                
                numeric_cols = df_display.select_dtypes(include=['float', 'int']).columns
                df_formatado = df_display.copy()
                for col in numeric_cols: df_formatado[col] = df_formatado[col].apply(formatar_moeda_br)

                tab1, tab2, tab3 = st.tabs(["📊 Resultado", "🚨 Divergências", "🕵️ Log"])
                with tab1:
                    st.dataframe(df_formatado, use_container_width=True)
                    st.download_button("Baixar Excel", to_excel(df_display), "conciliacao_final.xlsx")
                with tab2:
                    filtro = (df_final['Diferenca_Saldo_CC'].abs() > 0.01) | (df_final['Diferenca_Saldo_Aplic'].abs() > 0.01) | (df_final['Diferenca_Rendimento'].abs() > 0.01)
                    df_div = df_formatado[filtro]
                    if df_div.empty: st.info("Sem divergências!")
                    else: st.dataframe(df_div, use_container_width=True)
                with tab3: st.dataframe(df_log)
            else: st.error("Erro crítico: Tabela vazia.")
    else:
        st.warning("O botão só funciona se as caixas de status acima estiverem verdes ✅.")

