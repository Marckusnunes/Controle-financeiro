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
    page_title="Super Conciliador v2.0",
    layout="wide",
    page_icon="🔎",
    initial_sidebar_state="expanded"
)

# ==========================================
# 1. FUNÇÕES DE LIMPEZA E CHAVES (CORE)
# ==========================================

def gerar_chave_padronizada(texto_conta):
    """
    Padroniza a chave para cruzar os dados.
    Estratégia: Remove tudo que não é número e pega os últimos 7 dígitos.
    Ex: '7695-3' -> '0007695'
    Ex: '0000000076953' -> '0007695'
    """
    if isinstance(texto_conta, str):
        # Remove letras, traços, barras, espaços
        parte_numerica = re.sub(r'\D', '', texto_conta)
        if not parte_numerica: 
            return None
        # Pega os últimos 7 dígitos e garante zeros à esquerda
        ultimos_7_digitos = parte_numerica[-7:]
        return ultimos_7_digitos.zfill(7)
    return None

def limpar_valor_monetario(valor_str):
    """
    Transforma texto financeiro (ex: '1.500,20 D') em número (-1500.20).
    """
    if not isinstance(valor_str, str): return 0.0
    
    valor_upper = valor_str.upper()
    eh_negativo = 'D' in valor_upper or '-' in valor_str or 'DEB' in valor_upper
    
    # Mantém apenas dígitos, vírgula e ponto
    limpo = re.sub(r'[^\d,\.]', '', valor_str)
    
    try:
        if not limpo: return 0.0
        
        # Lógica para formato Brasileiro (1.000,00)
        if ',' in limpo and '.' in limpo:
             # Remove o ponto de milhar e troca vírgula decimal por ponto
             limpo = limpo.replace('.', '').replace(',', '.')
        elif ',' in limpo:
             # Apenas troca vírgula por ponto
             limpo = limpo.replace(',', '.')
        # Se só tiver ponto, assume que é inglês ou que o ponto é decimal se tiver 2 casas
        
        valor_float = float(limpo)
        return -valor_float if eh_negativo else valor_float
    except ValueError:
        return 0.0

# ==========================================
# 2. MOTOR DE LEITURA DE PDF (REFORMULADO)
# ==========================================

def extrair_pdf_melhorado(arquivo, tipo_extrato):
    """
    Tenta extrair Conta, Saldo e Rendimento de forma genérica e robusta.
    Serve tanto para BB quanto para Caixa.
    """
    try:
        doc = fitz.open(stream=arquivo.read(), filetype="pdf")
        texto_completo = ""
        # Lê todas as páginas
        for pag in doc:
            texto_completo += pag.get_text() + "\n"
        doc.close()
        
        # --- 1. EXTRAÇÃO DA CONTA (O mais crítico) ---
        conta_encontrada = "N/A"
        
        # Padrões de Regex para achar a conta (ordem de prioridade)
        padroes_conta = [
            r"Conta\s*Corrente\s*[:\s]*([\d\.\-\/]+)",      # Padrão Claro
            r"Conta\s*[:\s]*([\d\.\-\/]+)",                  # Padrão Simples
            r"Agência.*?Conta.*?([\d\.\-]{5,})",             # Padrão mesma linha
            r"Conta\s*Vinculada\s*[:\s]*([\d\.\-\/]+)",      # Padrão Caixa Gov
            r"(\d{4,}\s*\/\s*\d{3}\s*\/\s*\d{8,}-\d)",       # Padrão Caixa visual (Ag/Op/Conta)
            r"Cliente\s*[:\s]*.*?([\d\.\-]{5,})"             # Padrão BB antigo
        ]
        
        for p in padroes_conta:
            match = re.search(p, texto_completo, re.IGNORECASE)
            if match:
                # Limpa a conta encontrada para remover lixo
                conta_raw = match.group(1).strip()
                # Validação simples: tem que ter pelo menos 4 números
                if len(re.sub(r'\D', '', conta_raw)) > 4:
                    conta_encontrada = conta_raw
                    break
        
        # --- 2. EXTRAÇÃO DE VALORES (Saldo e Rendimento) ---
        saldo_final = 0.0
        rendimento_total = 0.0
        
        # Divide em linhas para analisar contexto
        linhas = texto_completo.split('\n')
        
        # Regex para capturar dinheiro (ex: 1.000,00 ou 50,00) no FIM da linha
        # Procura um número, vírgula, dois números e opcionalmente C/D ou espaço no fim
        regex_valor = r"(\d{1,3}(?:\.\d{3})*,\d{2})\s*([CD])?$" 
        
        for linha in linhas:
            linha_upper = linha.upper().strip()
            
            # -- Lógica de Saldo --
            # Palavras-chave que indicam o saldo final do extrato
            gatilhos_saldo = ["SALDO FINAL", "SALDO ATUAL", "SALDO EM", "SALDO TOTAL", "SALDO DISPONIVEL", "SALDO CREDOR", "SALDO DEVEDOR"]
            # Ignorar saldos parciais
            ignorar = ["ANTERIOR", "BLOQUEADO", "PROVISORIO", "LIMITE"]
            
            if any(g in linha_upper for g in gatilhos_saldo) and not any(i in linha_upper for i in ignorar):
                # Tenta pegar o valor na linha
                match_val = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linha)
                if match_val:
                    # Se achou, atualiza (o último válido do arquivo geralmente é o saldo final real)
                    valor_str = match_val.group(1)
                    # Verifica se tem D ou C na linha
                    sinal = "D" if " D" in linha_upper or "DEB" in linha_upper else "C"
                    saldo_final = limpar_valor_monetario(f"{valor_str} {sinal}")

            # Fallback: Se o extrato diz "SALDO", mas o valor está na linha de baixo (comum em PDF)
            # (Lógica simplificada: se a linha é só "SALDO", olha a próxima linha no futuro - complexo para fitz simples)
            
            # -- Lógica de Rendimento (Apenas se for extrato de investimento) --
            if tipo_extrato == 'INV':
                gatilhos_rend = ["RENDIMENTO BRUTO", "RENTABILIDADE", "RENDIMENTO NO MÊS", "RENDIMENTO LIQUIDO", "RENTAB. NO MES"]
                if any(g in linha_upper for g in gatilhos_rend) and "ACUMULADO" not in linha_upper:
                    match_val = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linha)
                    if match_val:
                        valor = limpar_valor_monetario(match_val.group(1))
                        # Filtro de segurança: Rendimento raramente é maior que 100 milhões num mês (evita pegar saldo total como rendimento)
                        if valor < 100000000:
                            rendimento_total += valor

        # Se não achou saldo nenhum, mas achou "Saldo Anterior" e o extrato é de conta sem movimento
        if saldo_final == 0.0 and ("NAO HOUVE MOVIMENTO" in texto_completo.upper() or "SEM MOVIMENTO" in texto_completo.upper()):
             match_ant = re.search(r"SALDO ANTERIOR.*?(\d{1,3}(?:\.\d{3})*,\d{2})", texto_completo, re.IGNORECASE | re.DOTALL)
             if match_ant:
                 saldo_final = limpar_valor_monetario(match_ant.group(1))

        return {
            "Conta": conta_encontrada,
            "Saldo": saldo_final,
            "Rendimento": rendimento_total,
            "Texto_Raw": texto_completo[:500] + "..." # Guarda um pedaço para debug
        }
        
    except Exception as e:
        return {"Conta": "Erro", "Saldo": 0.0, "Rendimento": 0.0, "Texto_Raw": str(e)}

# ==========================================
# 3. LEITURA CONTÁBIL
# ==========================================

def encontrar_coluna_chave(df):
    """Procura coluna de conta/domicílio."""
    candidatos = ['Domicílio bancário', 'Domicilio', 'Conta', 'Nº Conta', 'Descrição', 'Descricao']
    for col in df.columns:
        for cand in candidatos:
            if cand.lower() in str(col).lower(): return col
    return None

def encontrar_coluna_valor(df):
    """Procura coluna de valor."""
    candidatos = ['Saldo Final', 'Saldo Atual', 'Movimento', 'Valor']
    for col in df.columns:
        if 'anterior' in str(col).lower(): continue # Pula saldo anterior
        for cand in candidatos:
            if cand.lower() in str(col).lower(): return col
    return None

def processar_contabil(arquivo, tipo='SALDO'):
    """
    Lê CSV contábil.
    tipo='SALDO' -> Procura Saldo Final
    tipo='RENDIMENTO' -> Procura Movimento/Valor e soma
    """
    if arquivo is None: return pd.DataFrame()
    
    try:
        # Tenta header 1 (padrão Gov)
        df = pd.read_csv(arquivo, encoding='latin-1', sep=';', header=1, dtype=str)
        col_chave = encontrar_coluna_chave(df)
        
        # Se falhar, tenta header 0
        if not col_chave:
            arquivo.seek(0)
            df = pd.read_csv(arquivo, encoding='latin-1', sep=';', header=0, dtype=str)
            col_chave = encontrar_coluna_chave(df)
            
        if not col_chave:
            return pd.DataFrame() # Falhou silenciosamente para não quebrar tudo

        col_valor = encontrar_coluna_valor(df)
        if not col_valor: return pd.DataFrame()
        
        # Processa
        df['Chave Primaria'] = df[col_chave].apply(gerar_chave_padronizada)
        df = df.dropna(subset=['Chave Primaria'])
        
        df['Valor_Numerico'] = df[col_valor].astype(str).apply(limpar_valor_monetario)
        
        if tipo == 'SALDO':
            # Pivotar se possível (separar CC de Aplic)
            if 'Conta contábil' in df.columns:
                df_pivot = df.pivot_table(index='Chave Primaria', columns='Conta contábil', values='Valor_Numerico', aggfunc='sum').reset_index()
                
                # Tenta identificar colunas 1111101/901 (CC) e 1111150 (Aplic)
                col_mov = next((c for c in df_pivot.columns if '111111901' in str(c) or 'Conta Movimento' in str(c)), None)
                col_app = next((c for c in df_pivot.columns if '1111150' in str(c) or 'Aplicação' in str(c)), None)
                
                df_res = pd.DataFrame()
                df_res['Chave Primaria'] = df_pivot['Chave Primaria']
                df_res['Saldo_Contabil_CC'] = df_pivot[col_mov].fillna(0) if col_mov else 0.0
                df_res['Saldo_Contabil_Aplic'] = df_pivot[col_app].fillna(0) if col_app else 0.0
                
                # Recupera descrição
                desc = df[['Chave Primaria', col_chave]].drop_duplicates(subset='Chave Primaria')
                df_res = df_res.merge(desc, on='Chave Primaria', how='left')
                df_res.rename(columns={col_chave: 'Descrição'}, inplace=True)
                return df_res
            else:
                # Se não tiver conta contábil, assume tudo como CC
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
        st.error(f"Erro ao ler CSV Contábil ({tipo}): {e}")
        return pd.DataFrame()

# ==========================================
# 4. CONSOLIDAÇÃO
# ==========================================

def executar_processo(file_saldos, file_rendim, lista_extratos_cc, lista_extratos_inv):
    
    # 1. Contabilidade
    df_saldos = processar_contabil(file_saldos, 'SALDO')
    df_rendim = processar_contabil(file_rendim, 'RENDIMENTO')
    
    if df_saldos.empty:
        st.error("Erro no arquivo de Saldos. Verifique o cabeçalho.")
        return pd.DataFrame(), pd.DataFrame()
        
    # Junta contabilidade
    df_contabil = df_saldos
    if not df_rendim.empty:
        df_contabil = pd.merge(df_saldos, df_rendim, on='Chave Primaria', how='outer').fillna(0)
    else:
        df_contabil['Rendimento_Contabil'] = 0.0
        
    # 2. Extratos Bancários
    dados_banco = []
    log_leitura = [] # Para debug
    
    # Processa CC
    for f in lista_extratos_cc:
        res = extrair_pdf_melhorado(f, 'CC')
        chave = gerar_chave_padronizada(res['Conta'])
        dados_banco.append({
            'Chave Primaria': chave,
            'Saldo_Banco_CC': res['Saldo'],
            'Saldo_Banco_Aplic': 0.0,
            'Rendimento_Banco': 0.0
        })
        log_leitura.append({'Arquivo': f.name, 'Conta Lida': res['Conta'], 'Chave Gerada': chave, 'Saldo': res['Saldo'], 'Tipo': 'CC', 'Raw': res['Texto_Raw']})
        
    # Processa Investimentos
    for f in lista_extratos_inv:
        res = extrair_pdf_melhorado(f, 'INV')
        chave = gerar_chave_padronizada(res['Conta'])
        dados_banco.append({
            'Chave Primaria': chave,
            'Saldo_Banco_CC': 0.0,
            'Saldo_Banco_Aplic': res['Saldo'],
            'Rendimento_Banco': res['Rendimento']
        })
        log_leitura.append({'Arquivo': f.name, 'Conta Lida': res['Conta'], 'Chave Gerada': chave, 'Saldo': res['Saldo'], 'Rendimento': res['Rendimento'], 'Tipo': 'INV', 'Raw': res['Texto_Raw']})

    df_log = pd.DataFrame(log_leitura)
    
    if not dados_banco:
        return pd.DataFrame(), df_log
        
    df_banco_raw = pd.DataFrame(dados_banco)
    df_banco_raw = df_banco_raw.dropna(subset=['Chave Primaria'])
    
    # Consolida Banco (Soma por chave)
    df_banco = df_banco_raw.groupby('Chave Primaria').sum().reset_index()
    
    # 3. Cruzamento Final
    df_final = pd.merge(df_contabil, df_banco, on='Chave Primaria', how='inner')
    
    # Diferenças
    df_final['Diferenca_Saldo_CC'] = df_final['Saldo_Contabil_CC'] - df_final['Saldo_Banco_CC']
    df_final['Diferenca_Saldo_Aplic'] = df_final['Saldo_Contabil_Aplic'] - df_final['Saldo_Banco_Aplic']
    df_final['Diferenca_Rendimento'] = df_final['Rendimento_Contabil'] - df_final['Rendimento_Banco']
    
    # Seleciona colunas
    cols_desejadas = [
        'Descrição', 'Chave Primaria',
        'Saldo_Contabil_CC', 'Saldo_Banco_CC', 'Diferenca_Saldo_CC',
        'Saldo_Contabil_Aplic', 'Saldo_Banco_Aplic', 'Diferenca_Saldo_Aplic',
        'Rendimento_Contabil', 'Rendimento_Banco', 'Diferenca_Rendimento'
    ]
    # Filtra só as que existem
    cols_finais = [c for c in cols_desejadas if c in df_final.columns]
    
    return df_final[cols_finais], df_log

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Conciliacao')
        # Ajuste largura
        ws = writer.sheets['Conciliacao']
        for i, col in enumerate(df.columns):
            ws.column_dimensions[get_column_letter(i+1)].width = 20
    return output.getvalue()

# ==========================================
# 5. INTERFACE
# ==========================================

st.title("🔎 Super Conciliador de Contas")
st.markdown("### V2.0 - Leitura Otimizada")

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Contabilidade (CSV)")
    f_saldos = st.file_uploader("Relatório Saldos", type='csv')
    f_rendim = st.file_uploader("Relatório Rendimentos", type='csv')

with col2:
    st.subheader("2. Banco (PDFs)")
    f_cc = st.file_uploader("Extratos C/C (Todos os Bancos)", type='pdf', accept_multiple_files=True)
    f_inv = st.file_uploader("Extratos Aplicação (Todos os Bancos)", type='pdf', accept_multiple_files=True)

if st.button("Executar Conciliação", type="primary"):
    if f_saldos and (f_cc or f_inv):
        with st.spinner("Analisando documentos..."):
            
            # Combina listas (para simplificar, assumimos que o robô se vira com BB/Caixa misturado)
            df_res, df_log = executar_processo(f_saldos, f_rendim, f_cc if f_cc else [], f_inv if f_inv else [])
            
            if not df_res.empty:
                st.success("Processamento Finalizado!")
                
                tab1, tab2, tab3 = st.tabs(["📊 Resultado", "⚠️ Divergências", "🕵️ Raio-X (Debug)"])
                
                with tab1:
                    st.dataframe(df_res.style.format("{:,.2f}"))
                    st.download_button("Baixar Excel", to_excel(df_res), "conciliacao_v2.xlsx")
                
                with tab2:
                    div = df_res[
                        (df_res['Diferenca_Saldo_CC'].abs() > 0.01) |
                        (df_res['Diferenca_Saldo_Aplic'].abs() > 0.01) |
                        (df_res['Diferenca_Rendimento'].abs() > 0.01)
                    ]
                    if div.empty:
                        st.balloons()
                        st.info("Zero divergências encontradas!")
                    else:
                        st.error(f"{len(div)} contas com diferença.")
                        st.dataframe(div.style.format("{:,.2f}"))
                
                with tab3:
                    st.warning("Use esta aba para entender por que algum extrato não foi lido.")
                    st.write("Aqui mostramos o que o robô leu de cada arquivo PDF:")
                    st.dataframe(df_log)
                    
            else:
                st.error("Não houve cruzamento de dados.")
                st.write("Verifique a aba 'Raio-X' abaixo para ver se os PDFs foram lidos corretamente.")
                if not df_log.empty:
                    st.subheader("Log de Leitura dos PDFs")
                    st.dataframe(df_log)
    else:
        st.warning("Falta arquivo de Saldos ou Extratos.")
