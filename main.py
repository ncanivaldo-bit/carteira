import os
import json
import time
import gspread
import pandas as pd
import yfinance as yf
from datetime import datetime
from dateutil.relativedelta import relativedelta
from google.oauth2.service_account import Credentials

# ==============================================================================
# 1. AUTENTICAÇÃO E CONEXÃO
# ==============================================================================
print("Autenticando no Google Sheets...")

chave_json_str = os.environ.get('GCP_CREDENTIALS')
if not chave_json_str:
    raise ValueError("ERRO: A variável de ambiente 'GCP_CREDENTIALS' não foi encontrada.")

credenciais_dict = json.loads(chave_json_str)
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = Credentials.from_service_account_info(credenciais_dict, scopes=scopes)
gc = gspread.authorize(creds)

planilha = gc.open("base")
aba_mov = planilha.worksheet("Movimentacao")
aba_hist = planilha.worksheet("Hist_Precos")

conversao_tickers = {"MALL11": "PMLL11", "CVBI11": "PCIP11", "BOML": "BPML11"}
ticker_origem = {novo: antigo for antigo, novo in conversao_tickers.items()}

# ==============================================================================
# 2. CARREGAMENTO E CORREÇÃO DE DATAS DA MOVIMENTAÇÃO
# ==============================================================================
df_mov_records = pd.DataFrame(aba_mov.get_all_records(value_render_option='UNFORMATTED_VALUE'))

def converter_data(val):
    if pd.isna(val) or val == '':
        return pd.NaT
    if isinstance(val, (int, float)):
        try:
            return pd.to_datetime('1899-12-30') + pd.to_timedelta(val, unit='D')
        except (ValueError, TypeError):
            return pd.NaT
    try:
        return pd.to_datetime(val, dayfirst=True, errors='coerce')
    except (ValueError, TypeError):
        return pd.NaT

if 'Data' in df_mov_records.columns:
    df_mov_records['Data'] = df_mov_records['Data'].apply(converter_data)

# ==============================================================================
# 3. MOTOR DE ATUALIZAÇÃO HISTÓRICA E PLANILHA (YFINANCE EM LOTE)
# ==============================================================================
print("Atualizando histórico de preços (Yahoo Finance em lote)...")
if 'Ticker' in df_mov_records.columns:
    ativos_yf = set([conversao_tickers.get(str(tk).strip(), str(tk).strip()) for tk in df_mov_records['Ticker'].unique() if str(tk).strip() != ''])
else:
    ativos_yf = set()

df_hist = pd.DataFrame(aba_hist.get_all_records()) if aba_hist.get_all_values() and aba_hist.get_all_values()[0][0] == 'Chave_Merge' else pd.DataFrame()
mes_atual_str = datetime.now().strftime('%Y-%m')
novas_linhas = []

if ativos_yf:
    datas_inicio_por_ticker = {}
    min_data_inicio = datetime.now()

    for tk in ativos_yf:
        df_tk_hist = df_hist[df_hist['Ticker'] == tk] if not df_hist.empty and 'Ticker' in df_hist.columns else pd.DataFrame()
        data_inicio = datetime.strptime(df_tk_hist['Chave_Merge'].max(), '%Y-%m') + relativedelta(months=1) if not df_tk_hist.empty else datetime(2023, 5, 1)
        data_inicio = max(data_inicio, datetime(2023, 5, 1))
        data_inicio = data_inicio.replace(day=1)
        if data_inicio.strftime('%Y-%m') <= mes_atual_str:
            datas_inicio_por_ticker[tk] = data_inicio
            if data_inicio < min_data_inicio:
                min_data_inicio = data_inicio

    if datas_inicio_por_ticker:
        tickers_yf_lista = [f"{tk}.SA" for tk in datas_inicio_por_ticker.keys()]
        try:
            dados_yf = yf.download(tickers_yf_lista, start=min_data_inicio.strftime('%Y-%m-%d'), progress=False, auto_adjust=True)
            if not dados_yf.empty:
                for tk in datas_inicio_por_ticker.keys():
                    tk_yf = f"{tk}.SA"
                    data_limite = datas_inicio_por_ticker[tk]
                    try:
                        if len(tickers_yf_lista) == 1:
                            precos = dados_yf['Close']
                            if isinstance(precos, pd.DataFrame):
                                precos = precos.iloc[:, 0]
                        else:
                            if isinstance(dados_yf.columns, pd.MultiIndex):
                                if 'Close' in dados_yf.columns.get_level_values(0) and tk_yf in dados_yf.columns.get_level_values(1):
                                    precos = dados_yf['Close'][tk_yf]
                                elif 'Close' in dados_yf.columns.get_level_values(1) and tk_yf in dados_yf.columns.get_level_values(0):
                                    precos = dados_yf[tk_yf]['Close']
                                else:
                                    continue
                            else:
                                precos = dados_yf['Close'] if 'Close' in dados_yf.columns else pd.Series(dtype=float)

                        if not precos.empty:
                            precos = precos.dropna()
                            for data, preco in precos.resample('ME').last().dropna().items():
                                if data >= data_limite and preco > 0:
                                    novas_linhas.append([data.strftime('%Y-%m'), tk, round(float(preco), 2)])
                    except (ValueError, TypeError, KeyError) as e_tk:
                        print(f"Aviso: Não foi possível processar dados de {tk_yf}: {e_tk}")
        except Exception as e:
            print(f"Aviso: Erro ao baixar dados em lote do Yahoo Finance: {e}")

if novas_linhas:
    aba_hist.append_rows(novas_linhas)

# ==============================================================================
# 4. ATUALIZAÇÃO DO CADASTRO (INF_ATIVOS)
# ==============================================================================
print("Atualizando aba Inf_Ativos...")
try:
    aba_inf = planilha.worksheet("Inf_Ativos")
except:
    aba_inf = planilha.add_worksheet(title="Inf_Ativos", rows="1000", cols="20")

if not df_mov_records.empty and 'Ticker' in df_mov_records.columns:
    df_inf = df_mov_records.iloc[:, [0, 8, 9, 10, 11]].copy()
    nome_col_ticker = df_inf.columns[0]
    df_inf[nome_col_ticker] = df_inf[nome_col_ticker].astype(str).str.strip().replace(conversao_tickers)
    df_inf = df_inf[df_inf[nome_col_ticker] != ""].drop_duplicates(subset=[nome_col_ticker], keep='last').sort_values(by=nome_col_ticker)
    df_inf['Preço_Atual'] = df_inf[nome_col_ticker].apply(lambda tk: f'=IFERROR(GOOGLEFINANCE("BVMF:{tk}"; "price"); 0)')

    aba_inf.clear()
    aba_inf.update(range_name='A1', values=[df_inf.columns.values.tolist()] + df_inf.values.tolist(), value_input_option='USER_ENTERED')

# ==============================================================================
# 5. MOTOR INTELIGENTE DE CÁLCULO DE PREÇO MÉDIO
# ==============================================================================
print("Processando histórico e calculando Preços Médios...")

def garantir_inteiro(x):
    try:
        if isinstance(x, (int, float)): return int(x)
        return int(float(str(x).replace('R$', '').replace('.', '').replace(',', '.').strip()))
    except: return 0

def garantir_valor_financeiro(x):
    try:
        if isinstance(x, (int, float)): return float(x)
        return float(str(x).replace('R$', '').replace('.', '').replace(',', '.').strip())
    except: return 0.0

df_mov_calc = df_mov_records.copy()
df_mov_calc = df_mov_calc.sort_values(by=['Data', 'Movimentação']).reset_index(drop=True)

pm_dict = {}
historico_operacoes = []
ordem_execucao = 0

for idx, row in df_mov_calc.iterrows():
    tk = str(row.get('Ticker', '')).strip()
    if not tk: continue

    mov = str(row.get('Movimentação', '')).strip().title()
    es = str(row.get('Entrada/Saída', row.get('Entrada/Saida', ''))).strip().title()
    qtd = garantir_inteiro(row.get('Quantidade', 0))
    valor_op = garantir_valor_financeiro(row.get('Valor da Operação', 0))
    preco_unit = garantir_valor_financeiro(row.get('Preço unitário', 0))

    if valor_op == 0 and preco_unit > 0: valor_op = qtd * preco_unit
    if tk not in pm_dict: pm_dict[tk] = {'qtd': 0, 'custo_acumulado': 0.0, 'pm': 0.0}

    if mov == 'Compra' and es == 'Credito':
        pm_dict[tk]['qtd'] += qtd
        pm_dict[tk]['custo_acumulado'] += valor_op
        if pm_dict[tk]['qtd'] > 0: pm_dict[tk]['pm'] = pm_dict[tk]['custo_acumulado'] / pm_dict[tk]['qtd']

    elif mov in ['Venda', 'Atualização'] and es == 'Debito':
        pm_dict[tk]['qtd'] -= qtd
        if pm_dict[tk]['qtd'] <= 0:
            pm_dict[tk]['qtd'] = 0
            pm_dict[tk]['custo_acumulado'] = 0.0
        else:
            pm_dict[tk]['custo_acumulado'] = pm_dict[tk]['qtd'] * pm_dict[tk]['pm']

    elif mov == 'Desdobro' and es == 'Credito':
        pm_dict[tk]['qtd'] += qtd
        if pm_dict[tk]['qtd'] > 0: pm_dict[tk]['pm'] = pm_dict[tk]['custo_acumulado'] / pm_dict[tk]['qtd']

    elif mov == 'Atualização' and es == 'Credito':
        pm_dict[tk]['qtd'] += qtd
        if valor_op > 0:
            pm_dict[tk]['custo_acumulado'] += valor_op
        else:
            tk_antigo = ticker_origem.get(tk)
            if tk_antigo and tk_antigo in pm_dict:
                pm_dict[tk]['custo_acumulado'] += (qtd * pm_dict[tk_antigo]['pm'])
                pm_dict[tk_antigo]['qtd'] -= qtd
                if pm_dict[tk_antigo]['qtd'] <= 0:
                    pm_dict[tk_antigo]['qtd'] = 0
                    pm_dict[tk_antigo]['custo_acumulado'] = 0.0
                else:
                    pm_dict[tk_antigo]['custo_acumulado'] = pm_dict[tk_antigo]['qtd'] * pm_dict[tk_antigo]['pm']

                if pd.notna(row['Data']):
                    ordem_execucao += 1
                    historico_operacoes.append({'Ordem': ordem_execucao, 'Data': row['Data'], 'AnoMes': row['Data'].strftime('%Y-%m'), 'Ticker': tk_antigo, 'Qtd': pm_dict[tk_antigo]['qtd'], 'PM': pm_dict[tk_antigo]['pm'], 'Total_Investido': pm_dict[tk_antigo]['custo_acumulado']})

        if pm_dict[tk]['qtd'] > 0: pm_dict[tk]['pm'] = pm_dict[tk]['custo_acumulado'] / pm_dict[tk]['qtd']

    ordem_execucao += 1
    if pd.notna(row['Data']):
        historico_operacoes.append({'Ordem': ordem_execucao, 'Data': row['Data'], 'AnoMes': row['Data'].strftime('%Y-%m'), 'Ticker': tk, 'Qtd': pm_dict[tk]['qtd'], 'PM': pm_dict[tk]['pm'], 'Total_Investido': pm_dict[tk]['custo_acumulado']})

# ==============================================================================
# 6. CONSOLIDAÇÃO - SALDO GERAL (ABA CARTEIRA)
# ==============================================================================
resultados = []
for tk, dados in pm_dict.items():
    if tk not in conversao_tickers.keys() and dados['qtd'] > 0:
        resultados.append({'Ticker': tk, 'Qtd_Real': dados['qtd'], 'Preço_Médio': dados['pm'], 'Total_Investido': dados['custo_acumulado']})

df_resultado = pd.DataFrame(resultados)

if not df_resultado.empty:
    df_resultado = df_resultado.sort_values(by='Total_Investido', ascending=False).reset_index(drop=True)
    df_resultado['Qtd_Real'] = df_resultado['Qtd_Real'].astype(int)
    df_resultado['Preço_Médio'] = df_resultado['Preço_Médio'].astype(float).round(4)
    df_resultado['Total_Investido'] = df_resultado['Total_Investido'].astype(float).round(2)
    df_resultado['Preço_Mercado'] = df_resultado['Ticker'].apply(lambda tk: f'=IFERROR(GOOGLEFINANCE("BVMF:{tk}"; "price"); 0)')
    df_resultado['Patrimonio_Liquido'] = [f'=B{i+2}*E{i+2}' for i in range(len(df_resultado))]
    df_resultado['Variação'] = [f'=IFERROR((F{i+2}/D{i+2})-1; 0)' for i in range(len(df_resultado))]
    df_resultado['Classificacao'] = [f'=IFERROR(VLOOKUP(A{i+2}; Inf_Ativos!$A$2:$F; 2; FALSE); "")' for i in range(len(df_resultado))]
    df_resultado['Tipo'] = [f'=IFERROR(VLOOKUP(A{i+2}; Inf_Ativos!$A$2:$F; 3; FALSE); "")' for i in range(len(df_resultado))]
    df_resultado['Seguimento'] = [f'=IFERROR(VLOOKUP(A{i+2}; Inf_Ativos!$A$2:$F; 4; FALSE); "")' for i in range(len(df_resultado))]
    df_resultado['Gestora'] = [f'=IFERROR(VLOOKUP(A{i+2}; Inf_Ativos!$A$2:$F; 5; FALSE); "")' for i in range(len(df_resultado))]

    try: aba_carteira = planilha.worksheet("Carteira")
    except: aba_carteira = planilha.add_worksheet(title="Carteira", rows="1000", cols="20")
    aba_carteira.clear() # Limpa aba inteira
    aba_carteira.update(range_name='A1', values=[df_resultado.columns.values.tolist()] + df_resultado.values.tolist(), value_input_option='USER_ENTERED')
    print("-> Aba 'Carteira' atualizada!")

# ==============================================================================
# 7. CONSOLIDAÇÃO - HISTÓRICO MENSAL (ABA PM_MENSAL)
# ==============================================================================
if historico_operacoes:
    df_hist_op = pd.DataFrame(historico_operacoes)
    df_fechamento_mes = df_hist_op.sort_values('Ordem').groupby(['Ticker', 'AnoMes']).last().reset_index()

    mes_min = df_fechamento_mes['AnoMes'].min()
    meses_totais = pd.period_range(start=mes_min, end=mes_atual_str, freq='M').strftime('%Y-%m').tolist()

    linhas_completas = []
    for tk in df_fechamento_mes['Ticker'].unique():
        mes_inicial = df_fechamento_mes[df_fechamento_mes['Ticker'] == tk]['AnoMes'].min()
        meses_ativos = [m for m in meses_totais if m >= mes_inicial]
        for m in meses_ativos: linhas_completas.append({'Ticker': tk, 'AnoMes': m})

    df_grid = pd.DataFrame(linhas_completas)
    df_pm_mensal = pd.merge(df_grid, df_fechamento_mes[['Ticker', 'AnoMes', 'Qtd', 'PM', 'Total_Investido']], on=['Ticker', 'AnoMes'], how='left')
    df_pm_mensal = df_pm_mensal.sort_values(by=['Ticker', 'AnoMes'])
    df_pm_mensal['Qtd'] = df_pm_mensal.groupby('Ticker')['Qtd'].ffill()
    df_pm_mensal['PM'] = df_pm_mensal.groupby('Ticker')['PM'].ffill()
    df_pm_mensal['Total_Investido'] = df_pm_mensal.groupby('Ticker')['Total_Investido'].ffill()
    df_pm_mensal = df_pm_mensal[df_pm_mensal['Qtd'] > 0].copy()
    df_pm_mensal['Ticker'] = df_pm_mensal['Ticker'].astype(str).str.strip().replace(conversao_tickers)

    df_pm_mensal = df_pm_mensal.groupby(['AnoMes', 'Ticker'], as_index=False).agg({'Qtd': 'sum', 'Total_Investido': 'sum'})
    df_pm_mensal['PM'] = df_pm_mensal['Total_Investido'] / df_pm_mensal['Qtd']

    df_hist_procv = pd.DataFrame(aba_hist.get_all_records(value_render_option='UNFORMATTED_VALUE'))
    if not df_hist_procv.empty:
        colunas = df_hist_procv.columns
        df_procv = df_hist_procv[[colunas[0], colunas[1], colunas[2]]].copy()
        df_procv.columns = ['AnoMes', 'Ticker', 'Cotação']
        df_procv['Ticker'] = df_procv['Ticker'].astype(str).str.strip().replace(conversao_tickers)
        df_procv = df_procv.drop_duplicates(subset=['AnoMes', 'Ticker'], keep='last')
        df_pm_mensal = pd.merge(df_pm_mensal, df_procv, on=['AnoMes', 'Ticker'], how='left')
        df_pm_mensal['Cotação'] = pd.to_numeric(df_pm_mensal['Cotação'], errors='coerce').fillna(0.0)
    else:
        df_pm_mensal['Cotação'] = 0.0

    df_pm_mensal['Qtd'] = df_pm_mensal['Qtd'].astype(int)
    df_pm_mensal['Valor_Mercado'] = df_pm_mensal['Qtd'] * df_pm_mensal['Cotação']

    df_pm_mensal['PM'] = df_pm_mensal['PM'].astype(float).round(4)
    df_pm_mensal['Total_Investido'] = df_pm_mensal['Total_Investido'].astype(float).round(2)
    df_pm_mensal['Cotação'] = df_pm_mensal['Cotação'].astype(float).round(4)
    df_pm_mensal['Valor_Mercado'] = df_pm_mensal['Valor_Mercado'].astype(float).round(2)

    df_pm_mensal = df_pm_mensal.sort_values(by=['AnoMes', 'Ticker']).reset_index(drop=True)
    df_pm_mensal = df_pm_mensal[['AnoMes', 'Ticker', 'Qtd', 'PM', 'Cotação', 'Total_Investido', 'Valor_Mercado']]

    for i, row in df_pm_mensal.iterrows():
        if row['AnoMes'] == mes_atual_str:
            tk = row['Ticker']
            df_pm_mensal.at[i, 'Cotação'] = f'=IFERROR(GOOGLEFINANCE("BVMF:{tk}"; "price"); 0)'
            df_pm_mensal.at[i, 'Valor_Mercado'] = f'=C{i+2}*E{i+2}'

    df_pm_mensal['AnoMes'] = df_pm_mensal['AnoMes'].apply(lambda x: f"'{x}")

    try: aba_pm_mensal = planilha.worksheet("Pm_mensal")
    except: aba_pm_mensal = planilha.add_worksheet(title="Pm_mensal", rows="5000", cols="10")

    aba_pm_mensal.clear()
    aba_pm_mensal.update(range_name='A1', values=[df_pm_mensal.columns.values.tolist()] + df_pm_mensal.values.tolist(), value_input_option='USER_ENTERED')
    print("-> Aba 'Pm_mensal' atualizada!")

# ==============================================================================
# 8. LEITURA DA ABA DE MÉTRICAS / MTERICAS PARA REBALANCEAMENTO
# ==============================================================================
print("Buscando Alvos de Rebalanceamento (Aba Metricas)...")
try:
    try:
        aba_metricas = planilha.worksheet("Metricas")
    except gspread.exceptions.WorksheetNotFound:
        aba_metricas = planilha.worksheet("Mtericas")

    records_metricas = aba_metricas.get_all_records(value_render_option='UNFORMATTED_VALUE')
    dict_metricas = {}

    for row in records_metricas:
        keys = list(row.keys())
        if len(keys) >= 2:
            seg = str(row[keys[0]]).strip()
            pct_raw = row[keys[1]]

            try:
                pct_val = float(pct_raw)
                if 0 < pct_val <= 1.0:
                    pct_val *= 100
                dict_metricas[seg] = round(pct_val, 2)
            except (ValueError, TypeError):
                pass

    dados_metricas_js = json.dumps(dict_metricas)
except Exception as e:
    print(f"Aviso: Aba 'Metricas' não encontrada ou vazia. O Dashboard usará pesos iguais. ({e})")
    dados_metricas_js = "{}"

# ==============================================================================
# 9. GERAÇÃO DO DASHBOARD E INSIGHTS (COM O SEU MOTOR)
# ==============================================================================
print("Aguardando o Sheets processar as fórmulas (5s)...")
time.sleep(5)
print("Gerando o Dashboard...")

df_pm_mensal_original = pd.DataFrame(planilha.worksheet("Pm_mensal").get_all_records(value_render_option='UNFORMATTED_VALUE'))
df_pm_mensal_original['AnoMes'] = df_pm_mensal_original['AnoMes'].astype(str).str.replace("'", "")

df_grafico = df_pm_mensal_original.groupby('AnoMes', as_index=False).agg({'Total_Investido': 'sum', 'Valor_Mercado': 'sum'})

patrimonio_valores = pd.to_numeric(df_grafico['Valor_Mercado'], errors='coerce').fillna(0.0)
investido_valores = pd.to_numeric(df_grafico['Total_Investido'], errors='coerce').fillna(0.0)
patrimonio_atual = patrimonio_valores.iloc[-1] if not patrimonio_valores.empty else 0.0
investido_atual = investido_valores.iloc[-1] if not investido_valores.empty else 0.0

df_carteira_sheet = pd.DataFrame(planilha.worksheet("Carteira").get_all_records(value_render_option='UNFORMATTED_VALUE'))
df_pizza = df_carteira_sheet.copy()

def get_col(df, name_hints, default_val=0):
    for hint in name_hints:
        col = [c for c in df.columns if hint in str(c).lower()]
        if col: return df[col[0]]
    return default_val

df_pizza['ticker'] = get_col(df_pizza, ['ticker', 'ativo'], 'Outros')
df_pizza['tipo'] = df_pizza.get('Tipo', 'Outros')
df_pizza['seguimento'] = df_pizza.get('Seguimento', 'Outros')
df_pizza['gestora'] = df_pizza.get('Gestora', 'Outros')
df_pizza['Preço_Médio'] = get_col(df_pizza, ['médio', 'medio', 'pm'], 0)
df_pizza['Preço_Mercado'] = get_col(df_pizza, ['mercado', 'cotação', 'cotacao'], 0)
df_pizza['Variação'] = get_col(df_pizza, ['varia', 'var'], 0)

qtd_col = [c for c in df_pizza.columns if 'qtd' in str(c).lower()][0] if any('qtd' in str(c).lower() for c in df_pizza.columns) else df_pizza.columns[1]
df_pizza['Qtd'] = pd.to_numeric(df_pizza[qtd_col], errors='coerce').fillna(0)

col_patrimonio = [c for c in df_pizza.columns if 'patrimonio' in str(c).lower() or 'patrimônio' in str(c).lower() or 'liquido' in str(c).lower()]
if col_patrimonio:
    df_pizza['valor'] = pd.to_numeric(df_pizza[col_patrimonio[0]], errors='coerce').fillna(0)
else:
    preco_col = [c for c in df_pizza.columns if 'mercado' in str(c).lower()][0]
    df_pizza['valor'] = df_pizza['Qtd'] * pd.to_numeric(df_pizza[preco_col], errors='coerce').fillna(0)

mapa_tipos = dict(zip(df_pizza['ticker'].astype(str).str.strip(), df_pizza['tipo']))
mapa_seguimentos = dict(zip(df_pizza['ticker'].astype(str).str.strip(), df_pizza['seguimento']))

try:
    df_mov_prov = df_mov_records[df_mov_records['Movimentação'].astype(str).str.strip().str.title().isin(['Juros Sobre Capital Próprio', 'Rendimento', 'Jcp'])].copy()
    col_valor_op = [c for c in df_mov_prov.columns if 'valor' in str(c).lower()][0]
    df_mov_prov['Valor_Num'] = pd.to_numeric(df_mov_prov[col_valor_op], errors='coerce').fillna(0.0)

    cols_mov_lower = [str(c).lower() for c in df_mov_prov.columns]
    if 'ativo' in cols_mov_lower: col_ticker_mov = df_mov_prov.columns[cols_mov_lower.index('ativo')]
    elif 'ticker' in cols_mov_lower: col_ticker_mov = df_mov_prov.columns[cols_mov_lower.index('ticker')]
    else: col_ticker_mov = df_mov_prov.columns[0]

    df_mov_prov['ticker'] = df_mov_prov[col_ticker_mov].astype(str).str.strip()
    df_mov_prov['tipo'] = df_mov_prov['ticker'].map(mapa_tipos).fillna('Outros')
    df_mov_prov['seguimento'] = df_mov_prov['ticker'].map(mapa_seguimentos).fillna('Outros')

    if 'Data' in df_mov_prov.columns:
        df_mov_prov['Data_Dt'] = df_mov_prov['Data'].apply(converter_data)
        data_limite_12m = datetime.now() - relativedelta(years=1)
        df_mov_prov['is_12m'] = df_mov_prov['Data_Dt'] >= data_limite_12m
        df_mov_prov['AnoMes'] = df_mov_prov['Data_Dt'].dt.to_period('M').astype(str)
        
        ultimo_mes_prov = df_mov_prov['AnoMes'].max()
        ultimos_proventos = df_mov_prov[df_mov_prov['AnoMes'] == ultimo_mes_prov]['Valor_Num'].sum() if pd.notna(ultimo_mes_prov) else 0.0
    else:
        df_mov_prov['is_12m'] = True
        df_mov_prov['AnoMes'] = "N/D"
        ultimos_proventos = 0.0

    dividendos_totais = df_mov_prov['Valor_Num'].sum()
    dividendos_12m = df_mov_prov.loc[df_mov_prov['is_12m'], 'Valor_Num'].sum()
    df_export_prov = df_mov_prov[['AnoMes', 'Valor_Num', 'ticker', 'tipo', 'seguimento', 'is_12m']].rename(columns={'Valor_Num': 'valor'})
    dados_proventos_raw_js = df_export_prov.to_json(orient='records')
except Exception as e:
    dividendos_totais, dividendos_12m, ultimos_proventos = 0.0, 0.0, 0.0
    dados_proventos_raw_js = "[]"

ganho_capital = patrimonio_atual - investido_atual
lucro_total = ganho_capital + dividendos_totais
rentabilidade_total_pct = (lucro_total / investido_atual * 100) if investido_atual > 0 else 0

var_patrimonio_pct = (ganho_capital / investido_atual * 100) if investido_atual > 0 else 0
pct_ultimos_proventos = (ultimos_proventos / investido_atual * 100) if investido_atual > 0 else 0

if len(df_grafico) >= 12:
    inv_12m_atras = pd.to_numeric(df_grafico.iloc[-12]['Total_Investido'], errors='coerce')
    pat_12m_atras = pd.to_numeric(df_grafico.iloc[-12]['Valor_Mercado'], errors='coerce')
    rentabilidade_12m_pct = ((patrimonio_atual - pat_12m_atras + dividendos_12m) / inv_12m_atras * 100) if inv_12m_atras > 0 else rentabilidade_total_pct
else:
    rentabilidade_12m_pct = rentabilidade_total_pct

# ----------------- INSIGHTS DE IA -----------------
insights_ia = []
if ganho_capital > 0:
    insights_ia.append(f"🟢 Seu patrimônio teve uma valorização de capital de <b>R$ {ganho_capital:,.2f}</b>.")
else:
    insights_ia.append(f"🔴 Sua carteira está descontada em <b>R$ {abs(ganho_capital):,.2f}</b>. Boa oportunidade para aportes?")

if investido_atual > 0:
    yoc = (dividendos_totais / investido_atual) * 100
    insights_ia.append(f"🔥 Seu <b>Yield on Cost (Histórico)</b> está em <b>{yoc:.2f}%</b>. Esse é o rendimento real do seu dinheiro!")

try:
    idx_max = pd.to_numeric(df_pizza['Variação'], errors='coerce').idxmax()
    if pd.notna(idx_max):
        campeao_tk = df_pizza.loc[idx_max, 'ticker']
        campeao_var = df_pizza.loc[idx_max, 'Variação'] * 100
        if campeao_var > 0:
            insights_ia.append(f"🏆 O ativo campeão da carteira hoje é <b>{campeao_tk}</b>, com <b>{campeao_var:.2f}%</b> de lucro.")
except:
    pass

html_insights = "".join([f"<li style='margin-bottom: 6px;'>{ins}</li>" for ins in insights_ia])

# Sinais e Cores
cor_lucro = "text-green" if lucro_total >= 0 else "text-red"
cor_ganho = "text-green" if ganho_capital >= 0 else "text-red"
cor_rent_12m = "text-green" if rentabilidade_12m_pct >= 0 else "text-red"
cor_rent_total = "text-green" if rentabilidade_total_pct >= 0 else "text-red"
cor_var_pat = "text-green" if var_patrimonio_pct >= 0 else "text-red"

sinal_rent_12m = "▲" if rentabilidade_12m_pct >= 0 else "▼"
sinal_rent_total = "▲" if rentabilidade_total_pct >= 0 else "▼"
sinal_var_pat = "▲" if var_patrimonio_pct >= 0 else "▼"

def formata_moeda(valor):
    sinal = "-" if valor < 0 else ""
    return f"{sinal}R$ {abs(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

str_patrimonio = formata_moeda(patrimonio_atual)
str_investido = formata_moeda(investido_atual)
str_lucro = formata_moeda(lucro_total)
str_ganho_capital = formata_moeda(ganho_capital)
str_dividendos_total = formata_moeda(dividendos_totais)
str_dividendos_12m = formata_moeda(dividendos_12m)
str_ultimos_proventos = formata_moeda(ultimos_proventos)
str_pct_ultimos_proventos = f"{pct_ultimos_proventos:.2f}%".replace(".", ",")
str_rentabilidade_total = f"{rentabilidade_total_pct:.2f}%".replace(".", ",")
str_rentabilidade_12m = f"{rentabilidade_12m_pct:.2f}%".replace(".", ",")
str_var_patrimonio = f"{sinal_var_pat} {abs(var_patrimonio_pct):.2f}%".replace(".", ",")

meses_js = json.dumps(df_grafico['AnoMes'].tolist())
valor_aplicado_js = json.dumps(pd.to_numeric(df_grafico['Total_Investido'], errors='coerce').fillna(0.0).tolist())
valor_mercado_js = json.dumps(pd.to_numeric(df_grafico['Valor_Mercado'], errors='coerce').fillna(0.0).tolist())
dados_carteira_js = df_pizza[['ticker', 'tipo', 'seguimento', 'gestora', 'valor', 'Preço_Médio', 'Preço_Mercado', 'Variação', 'Qtd']].fillna(0).to_json(orient='records')

html_template = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>Dashboard de Investimentos</title>
    <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        :root { 
            --bg-color: #0f172a; 
            --card-bg: #1e293b; 
            --text-main: #f8fafc; 
            --text-muted: #94a3b8; 
            --green: #10b981;
            --red: #ef4444; 
            --border-color: rgba(255, 255, 255, 0.08); 
            --primary: #3b82f6;
        }
        body { background-color: var(--bg-color); color: var(--text-main); font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; margin: 0; padding: 15px 20px; box-sizing: border-box; height: 100vh; overflow: hidden; display: flex; flex-direction: column; }
        header { display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--border-color); padding-bottom: 10px; margin-bottom: 15px; flex-shrink: 0; }
        .nav-links { display: flex; gap: 20px; }
        .nav-links a { color: var(--text-muted); text-decoration: none; font-size: 14.5px; font-weight: 500; cursor: pointer; padding-bottom: 10px; transition: color 0.2s;}
        .nav-links a:hover, .nav-links a.active { color: var(--primary); font-weight: 600;}
        .nav-links a.active { border-bottom: 2px solid var(--primary); }
        .header-buttons button { background-color: var(--primary); color: white; font-weight: bold; border: none; padding: 8px 16px; border-radius: 6px; cursor: pointer; font-size: 13px; transition: 0.2s; }
        .header-buttons button:hover { background-color: #2563eb; }
        .tab-content { display: none; flex-grow: 1; flex-direction: column; overflow-y: auto; padding-right: 5px; }
        .tab-content.active { display: flex; }
        .dashboard-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 15px; flex-shrink: 0; }
        .card { background-color: var(--card-bg); border: 1px solid var(--border-color); border-radius: 8px; padding: 16px 20px; display: flex; flex-direction: column; justify-content: space-between; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.2);}
        .kpi-title { font-size: 13.5px; color: var(--text-muted); font-weight: 500; margin-bottom: 6px; }
        .kpi-value { font-size: 22px; font-weight: 700; margin-bottom: 2px; display: flex; align-items: center; }
        .kpi-sub { font-size: 12.5px; color: var(--text-muted); display: flex; justify-content: space-between; margin-top: 10px;}
        .text-green { color: var(--green); } .text-red { color: var(--red); }
        .charts-grid { display: grid; grid-template-columns: 1.2fr 1fr; gap: 15px; min-height: 320px; flex-shrink: 0; margin-bottom: 15px; }
        .charts-grid-full { grid-template-columns: 1fr; min-height: 380px; margin-bottom: 5px;}
        .chart-container { width: 100%; height: 100%; min-height: 200px; position: relative; }
        .chart-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 5px; }
        .chart-title { font-size: 15.5px; font-weight: 600; color: var(--text-main); }
        .select-filter { background-color: var(--card-bg); border: 1px solid var(--border-color); color: var(--text-main); padding: 6px 12px; border-radius: 6px; font-size: 12px; outline: none; cursor: pointer; }
        .filtros-prov { display: flex; gap: 10px; }
        
        /* Nova Tabela de Rebalanceamento com Cabeçalhos Clicáveis e Coluna Congelada */
        .table-container { overflow-x: auto; max-width: 100%; position: relative; margin-top: 10px; flex-grow: 1; }
        table { width: 100%; border-collapse: separate; border-spacing: 0; text-align: left; font-size: 13px; min-width: 1200px; }
        th, td { padding: 12px 10px; border-bottom: 1px solid var(--border-color); white-space: nowrap; }
        th { color: var(--text-muted); font-weight: 600; position: sticky; top: 0; background-color: var(--card-bg); z-index: 2; cursor: pointer; user-select: none; transition: background 0.2s;}
        th:hover { background-color: #334155; color: var(--text-main); }
        th i { margin-left: 5px; font-size: 11px; }
        
        /* Congelando a primeira coluna (Ativo) */
        td:first-child, th:first-child { position: sticky; left: 0; z-index: 3; background-color: var(--card-bg); border-right: 1px solid var(--border-color); box-shadow: 2px 0 5px rgba(0,0,0,0.2);}
        th:first-child { z-index: 4; } /* O cabeçalho da primeira coluna precisa ficar acima da linha e da coluna */
        tr:hover td { background-color: #334155; }
        tr:hover td:first-child { background-color: #334155; } /* Mantém o hover na coluna congelada */

        @media (max-width: 992px) { body { height: auto; overflow-y: auto; overflow-x: hidden; } .dashboard-grid { grid-template-columns: repeat(2, 1fr); } .charts-grid { grid-template-columns: 1fr; display: flex; flex-direction: column; min-height: auto;} .chart-container { height: 350px; min-height: 350px; } }
        @media (max-width: 600px) { header { flex-direction: column; align-items: stretch; gap: 15px; padding-bottom: 15px; } .nav-links { justify-content: center; width: 100%; border-bottom: 1px solid var(--border-color); padding-bottom: 5px; } .header-buttons button { width: 100%; } .dashboard-grid { grid-template-columns: 1fr; } .chart-header { flex-direction: column; align-items: flex-start; gap: 10px; } .filtros-prov { flex-direction: column; width: 100%; } .select-filter { width: 100%; } .chart-container { height: 300px; min-height: 300px; } }
    </style>
</head>
<body>
    <header>
        <div class="nav-links">
            <a onclick="mudarAba('resumo')" id="link-resumo" class="active">Resumo</a>
            <a onclick="mudarAba('proventos')" id="link-proventos">Proventos</a>
            <a onclick="mudarAba('gestoras')" id="link-gestoras">Gestoras</a>
            <a onclick="mudarAba('rebalanceamento')" id="link-rebalanceamento">Rebalanceamento</a>
        </div>
    </header>

    <!-- ================= ABA 1: RESUMO ================= -->
    <div id="tab-resumo" class="tab-content active">
        
        <div class="card" style="background: linear-gradient(145deg, rgba(59,130,246,0.15) 0%, rgba(139,92,246,0.15) 100%); border-left: 4px solid var(--primary); margin-bottom: 15px;">
            <div class="kpi-title" style="color: var(--text-main); font-size: 15px; font-weight: 600;"><i class="fa-solid fa-wand-magic-sparkles" style="color: var(--primary); margin-right: 5px;"></i> Insights da Carteira</div>
            <ul id="ai-insights" style="margin-top: 10px; color: var(--text-muted); font-size: 13.5px; padding-left: 20px;">
                __INSIGHTS_HTML__
            </ul>
        </div>

        <div class="dashboard-grid">
            <div class="card"><div class="kpi-title">💰 Patrimônio total</div>
                <div class="kpi-value" style="color: var(--primary);">__PATRIMONIO__ <span class="__COR_VAR_PAT__" style="font-size: 14.5px; font-weight: 600; margin-left: 10px;">(__VAR_PATRIMONIO__)</span></div>
                <div class="kpi-sub"><span>Valor investido</span><span style="font-weight: 600; color: var(--text-main);">__INVESTIDO__</span></div>
            </div>
            <div class="card"><div class="kpi-title">💲 Lucro total</div><div class="kpi-value __COR_LUCRO__">__LUCRO__</div><div class="kpi-sub"><div style="display: flex; flex-direction: column; gap: 2px;"><span>Ganho</span><span class="__COR_GANHO__" style="font-size: 13px; font-weight: 600;">__GANHO_CAPITAL__</span></div><div style="display: flex; flex-direction: column; gap: 2px; text-align: right;"><span>Divs</span><span style="color: var(--text-main); font-size: 13px; font-weight: 600;">__DIVIDENDOS_TOTAL__</span></div></div></div>
            <div class="card"><div class="kpi-title">🪙 Últimos Proventos</div>
                <div class="kpi-value" style="color: var(--primary);">__ULTIMOS_PROVENTOS__ <span style="font-size: 14.5px; font-weight: 600; margin-left: 10px; color: var(--text-muted);">(__PCT_ULTIMOS_PROVENTOS__)</span></div>
                <div class="kpi-sub"><div style="display: flex; flex-direction: column; gap: 2px;"><span style="font-size: 11.5px; color: var(--text-muted);">Total Acumulado</span><span style="color: var(--text-main); font-size: 13px; font-weight: 600;">__DIVIDENDOS_TOTAL__</span></div></div>
            </div>
            <div class="card"><div class="kpi-title">📊 Rentabilidade (12M)</div><div style="display: flex; justify-content: space-between; align-items: center; margin-top: 6px;"><div><div class="kpi-value __COR_RENT_12M__" style="font-size: 19px; margin-bottom: 0;">__RENTABILIDADE_12M__ __SINAL_RENT_12M__</div></div><div style="border-left: 1px solid var(--border-color); padding-left: 15px;"><div style="font-size: 11px; color: var(--text-muted); margin-bottom: 2px;">Total</div><div class="__COR_RENT_TOTAL__" style="font-size: 16px; font-weight: bold;">__RENTABILIDADE_TOTAL__ __SINAL_RENT_TOTAL__</div></div></div></div>
        </div>

        <div class="charts-grid" style="grid-template-columns: 1fr; min-height: 250px;">
            <div class="card" style="display: flex; flex-direction: column;"><div class="chart-header"><div class="chart-title">Evolução do Patrimônio</div></div><div id="grafico-linha" class="chart-container"></div></div>
        </div>

        <div class="charts-grid">
            <div class="card" style="display: flex; flex-direction: column;">
                <div class="chart-header"><div class="chart-title">Jornada de Retorno Oficial (Waterfall)</div></div>
                <div id="grafico-waterfall" class="chart-container"></div>
            </div>
            <div class="card" style="display: flex; flex-direction: column;">
                <div class="chart-header"><div class="chart-title">Composição da Carteira (Ativos)</div></div>
                <div id="grafico-sunburst" class="chart-container"></div>
            </div>
        </div>
    </div>

    <!-- ================= ABA 2: PROVENTOS ================= -->
    <div id="tab-proventos" class="tab-content" style="overflow: hidden; padding-bottom: 5px;">
        <div class="dashboard-grid">
            <div class="card"><div class="kpi-title">💰 Últimos Dividendos</div><div class="kpi-value text-green" id="kpi-prov-ultimos">R$ 0,00</div><div class="kpi-sub"><span id="kpi-prov-ultimos-pct">0,00% sobre investido</span></div></div>
            <div class="card"><div class="kpi-title">💵 Proventos Filtrados</div><div class="kpi-value text-green" id="kpi-prov-total">R$ 0,00</div><div class="kpi-sub"><span>Acumulado histórico</span></div></div>
            <div class="card"><div class="kpi-title">📅 Proventos (12M)</div><div class="kpi-value text-green" id="kpi-prov-12m">R$ 0,00</div><div class="kpi-sub"><span>Últimos 12 Meses</span></div></div>
            <div class="card"><div class="kpi-title">📈 Média Mensal</div><div class="kpi-value" style="color: var(--primary);" id="kpi-prov-media">R$ 0,00</div><div class="kpi-sub"><span>Média por mês ativo</span></div></div>
        </div>
        <div class="charts-grid charts-grid-full" style="flex-grow: 1; min-height: 0; display: flex; flex-direction: column; margin-bottom: 0;">
            <div class="card" style="display: flex; flex-direction: column; flex-grow: 1;">
                <div class="chart-header"><div class="chart-title">Evolução Mensal de Proventos</div><div class="filtros-prov"><select id="f-tipo" class="select-filter" onchange="atualizarAbaProventos()"></select><select id="f-seguimento" class="select-filter" onchange="atualizarAbaProventos()"></select><select id="f-ativo" class="select-filter" onchange="atualizarAbaProventos()"></select></div></div>
                <div id="grafico-proventos-mensal" class="chart-container" style="flex-grow: 1; min-height: 0;"></div>
            </div>
        </div>
    </div>

    <!-- ================= ABA 3: GESTORAS ================= -->
    <div id="tab-gestoras" class="tab-content">
        <div class="charts-grid charts-grid-full" style="flex-grow: 1;">
            <div class="card" style="display: flex; flex-direction: column;">
                <div class="chart-header"><div class="chart-title">Alocação por Gestora</div></div>
                <div id="grafico-treemap" class="chart-container" style="min-height: 400px;"></div>
            </div>
        </div>
    </div>

    <!-- ================= ABA 4: REBALANCEAMENTO ================= -->
    <div id="tab-rebalanceamento" class="tab-content">
        <div class="dashboard-grid" style="grid-template-columns: 1fr;">
            <div class="card">
                <div class="kpi-title" style="font-size: 15.5px; font-weight: 600; color: var(--text-main);">⚖️ Estratégia de Alocação por Seguimento</div>
                <div class="kpi-sub" style="margin-top: 5px;">Os percentuais foram importados da aba Metricas da planilha. Ajuste aqui caso queira simular outro cenário.</div>
                <div id="segment-inputs" style="display: flex; gap: 15px; flex-wrap: wrap; margin-top: 15px;"></div>
                <div id="soma-alerta" style="font-size: 13.5px; font-weight: 600; margin-top: 15px;">Soma atual: 100%</div>
            </div>
        </div>

        <div class="card" style="flex-grow: 1; display: flex; flex-direction: column; min-height: 300px;">
            <div class="chart-header">
                <div class="chart-title">Radar de Aportes (Clique nos títulos para ordenar)</div>
            </div>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th onclick="sortTable('ticker')">Ativo <i class="fa-solid fa-sort"></i></th>
                            <th onclick="sortTable('tipo')">Tipo <i class="fa-solid fa-sort"></i></th>
                            <th onclick="sortTable('seg')">Seguimento <i class="fa-solid fa-sort"></i></th>
                            <th onclick="sortTable('gestora')">Gestora <i class="fa-solid fa-sort"></i></th>
                            <th onclick="sortTable('qtd')">Qtd <i class="fa-solid fa-sort"></i></th>
                            <th onclick="sortTable('pm')">PM <i class="fa-solid fa-sort"></i></th>
                            <th onclick="sortTable('preco')">Cotação <i class="fa-solid fa-sort"></i></th>
                            <th onclick="sortTable('var')">Var (%) <i class="fa-solid fa-sort"></i></th>
                            <th onclick="sortTable('current')">Valor Atual <i class="fa-solid fa-sort"></i></th>
                            <th onclick="sortTable('pct_atual')">% Carteira <i class="fa-solid fa-sort"></i></th>
                            <th onclick="sortTable('target_pct')">% Alvo <i class="fa-solid fa-sort"></i></th>
                            <th onclick="sortTable('target')">Valor Alvo <i class="fa-solid fa-sort"></i></th>
                            <th onclick="sortTable('diff')">Falta/Sobra (R$) <i class="fa-solid fa-sort"></i></th>
                        </tr>
                    </thead>
                    <tbody id="tabela-rebalanceamento"></tbody>
                </table>
            </div>
        </div>
    </div>

    <script>
        const configPlotly = { displayModeBar: false, responsive: true };
        
        const baseLayout = { 
            paper_bgcolor: 'rgba(0,0,0,0)', 
            plot_bgcolor: 'rgba(0,0,0,0)', 
            font: { color: '#94a3b8', size: 11.5, family: 'Segoe UI' }, 
            margin: { t: 30, l: 45, r: 15, b: 35 }, 
            hovermode: 'x unified', 
            xaxis: { showgrid: false, zeroline: false, fixedrange: true }, 
            yaxis: { showgrid: false, zeroline: false, fixedrange: true }, 
            colorway: ['#3b82f6', '#10b981', '#8b5cf6', '#f59e0b', '#ec4899', '#f97316', '#14b8a6'], 
            autosize: true, 
            dragmode: false 
        };
        const cloneLayout = (obj) => JSON.parse(JSON.stringify(obj));
        const formataMoedaJS = (valor) => valor.toLocaleString('pt-BR', {style: 'currency', currency: 'BRL'});
        const valInvestidoAtual = __INVESTIDO_ATUAL_NUM__;

        function mudarAba(nomeAba) {
            document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
            document.querySelectorAll('.nav-links a').forEach(el => el.classList.remove('active'));
            document.getElementById('tab-' + nomeAba).classList.add('active');
            document.getElementById('link-' + nomeAba).classList.add('active');
            window.dispatchEvent(new Event('resize'));
        }

        // ================= GRÁFICO LINHA =================
        let layoutLinha = cloneLayout(baseLayout); layoutLinha.showlegend = true; layoutLinha.legend = { orientation: 'h', y: 1.15, x: 0.1, font: {size: 11.5} };
        Plotly.newPlot('grafico-linha', [
            { x: __MESES_JS__, y: __VALOR_APLICADO_JS__, name: 'Valor Investido', type: 'scatter', mode: 'lines', line: { color: '#10b981', width: 3 }, marker: { size: 5 }}, 
            { x: __MESES_JS__, y: __VALOR_MERCADO_JS__, name: 'Valor de Mercado', type: 'scatter', mode: 'lines', line: { color: '#3b82f6', width: 3 }, marker: { size: 5 }}
        ], layoutLinha, configPlotly);

        // ================= GRÁFICO WATERFALL REVISADO =================
        // Investido -> Valorização -> Patrimônio -> Dividendos -> Riqueza Total
        const valInvestido = valInvestidoAtual;
        const valGanho = __GANHO_CAPITAL_NUM__;
        const valPatrimonio = valInvestido + valGanho;
        const valProv = __DIVIDENDOS_TOTAIS_NUM__;
        const valFinal = valPatrimonio + valProv;

        const traceWaterfall = {
            type: "waterfall", orientation: "v",
            measure: ["absolute", "relative", "total", "relative", "total"],
            x: ["Investido", "Oscilação Cotas", "Patrimônio", "Proventos", "Riqueza Acumulada"],
            y: [valInvestido, valGanho, valPatrimonio, valProv, valFinal],
            textposition: "outside",
            text: [formataMoedaJS(valInvestido), formataMoedaJS(valGanho), formataMoedaJS(valPatrimonio), formataMoedaJS(valProv), formataMoedaJS(valFinal)],
            decreasing: { marker: { color: "#ef4444" } },
            increasing: { marker: { color: "#10b981" } },
            totals: { marker: { color: "#3b82f6" } }
        };
        let layoutWater = cloneLayout(baseLayout);
        layoutWater.margin.t = 45; // Mais espaço pra cima
        Plotly.newPlot('grafico-waterfall', [traceWaterfall], layoutWater, configPlotly);

        // VARIÁVEIS COMPARTILHADAS (CARTEIRA)
        const dadosCarteira = __DADOS_CARTEIRA_JS__;
        const valPatrimonioAtual = __PATRIMONIO_ATUAL_NUM__;
        let carteiraTotal = 0;

        // ================= GRÁFICO SUNBURST =================
        let ids = ["Carteira"], labels = ["Carteira"], parents = [""], values = [0], totaisTipo = {}, totaisSeg = {};
        dadosCarteira.forEach(a => { let tipo = a.tipo || 'Outros', seg = a.seguimento || 'Outros', segId = tipo + " - " + seg; carteiraTotal += a.valor; totaisTipo[tipo] = (totaisTipo[tipo] || 0) + a.valor; if(!totaisSeg[segId]) totaisSeg[segId] = { label: seg, parent: tipo, val: 0 }; totaisSeg[segId].val += a.valor; });
        values[0] = carteiraTotal;
        for (let t in totaisTipo) { ids.push(t); labels.push(t); parents.push("Carteira"); values.push(totaisTipo[t]); }
        for (let s in totaisSeg) { ids.push(s); labels.push(totaisSeg[s].label); parents.push(totaisSeg[s].parent); values.push(totaisSeg[s].val); }
        dadosCarteira.forEach(a => { let tipo = a.tipo || 'Outros', seg = a.seguimento || 'Outros', segId = tipo + " - " + seg; ids.push(a.ticker); labels.push(a.ticker); parents.push(segId); values.push(a.valor); });

        let customText = [], customHover = [];
        for (let i = 0; i < values.length; i++) {
            let strMoeda = formataMoedaJS(values[i]);
            if (i === 0) { let strMoedaPat = formataMoedaJS(valPatrimonioAtual); customText.push(`<b>${labels[i]}</b><br>${strMoedaPat}`); customHover.push(`<b>${labels[i]}</b><br>Patrimônio Total: ${strMoedaPat}`); }
            else { let pct = carteiraTotal > 0 ? (values[i] / carteiraTotal) * 100 : 0; let strPct = pct.toFixed(2).replace('.', ',') + '%'; customText.push(`${labels[i]}<br>${strPct}`); customHover.push(`<b>${labels[i]}</b><br>Valor: ${strMoeda}<br>Participação: ${strPct}`); }
        }
        let layoutSunburst = cloneLayout(baseLayout); layoutSunburst.margin = { t: 10, l: 5, r: 5, b: 5 }; delete layoutSunburst.xaxis; delete layoutSunburst.yaxis;
        Plotly.newPlot('grafico-sunburst', [{ type: "sunburst", ids: ids, labels: labels, parents: parents, values: values, branchvalues: "total", text: customText, hovertext: customHover, textinfo: "text", hoverinfo: "text", marker: { line: { width: 1.5, color: '#1e293b' } } }], layoutSunburst, configPlotly);

        // ================= GRÁFICO TREEMAP (GESTORAS) =================
        let tmIds = ["Alocação"], tmLabels = ["Alocação"], tmParents = [""], tmValues = [carteiraTotal], totaisGestora = {};
        dadosCarteira.forEach(a => { let g = a.gestora || 'Outros'; totaisGestora[g] = (totaisGestora[g] || 0) + a.valor; });
        for (let g in totaisGestora) { tmIds.push(g); tmLabels.push(g); tmParents.push("Alocação"); tmValues.push(totaisGestora[g]); }
        dadosCarteira.forEach(a => { let g = a.gestora || 'Outros'; tmIds.push(a.ticker + " (" + g + ")"); tmLabels.push(a.ticker); tmParents.push(g); tmValues.push(a.valor); });

        let tmText = [], tmHover = [];
        for (let i = 0; i < tmValues.length; i++) {
            let strMoeda = formataMoedaJS(tmValues[i]);
            if (i === 0) {
                let strMoedaPat = formataMoedaJS(valPatrimonioAtual);
                tmText.push(`<b>${tmLabels[i]}</b><br>${strMoedaPat}`); tmHover.push(`<b>${tmLabels[i]}</b><br>Patrimônio Total: ${strMoedaPat}`);
            } else {
                let pct = carteiraTotal > 0 ? (tmValues[i] / carteiraTotal) * 100 : 0;
                let strPct = pct.toFixed(2).replace('.', ',') + '%';
                tmText.push(`<b>${tmLabels[i]}</b><br>${strPct}`); tmHover.push(`<b>${tmLabels[i]}</b><br>Valor: ${strMoeda}<br>Participação: ${strPct}`);
            }
        }
        let layoutTreemap = cloneLayout(baseLayout); layoutTreemap.margin = { t: 5, l: 5, r: 5, b: 5 }; delete layoutTreemap.xaxis; delete layoutTreemap.yaxis;
        Plotly.newPlot('grafico-treemap', [{ type: "treemap", ids: tmIds, labels: tmLabels, parents: tmParents, values: tmValues, branchvalues: "total", text: tmText, hovertext: tmHover, textinfo: "text", hoverinfo: "text", marker: { line: { width: 1.5, color: '#1e293b' } } }], layoutTreemap, configPlotly);

        // ================= LÓGICA DE PROVENTOS =================
        const dadosProvRaw = __DADOS_PROVENTOS_RAW_JS__;
        function popularFiltros() {
            let tipos = new Set(), segs = new Set(), ativos = new Set();
            dadosProvRaw.forEach(d => { tipos.add(d.tipo); segs.add(d.seguimento); ativos.add(d.ticker); });
            const addOp = (sel, val, text) => { sel.add(new Option(text, val)); };
            addOp(document.getElementById('f-tipo'), 'Todos', 'Tipo: Todos'); Array.from(tipos).sort().forEach(v => addOp(document.getElementById('f-tipo'), v, v));
            addOp(document.getElementById('f-seguimento'), 'Todos', 'Seguimento: Todos'); Array.from(segs).sort().forEach(v => addOp(document.getElementById('f-seguimento'), v, v));
            addOp(document.getElementById('f-ativo'), 'Todos', 'Ativo: Todos'); Array.from(ativos).sort().forEach(v => addOp(document.getElementById('f-ativo'), v, v));
        }
        function atualizarAbaProventos() {
            const vTipo = document.getElementById('f-tipo').value, vSeg = document.getElementById('f-seguimento').value, vAtivo = document.getElementById('f-ativo').value;
            let df = dadosProvRaw.filter(d => (vTipo === 'Todos' || d.tipo === vTipo) && (vSeg === 'Todos' || d.seguimento === vSeg) && (vAtivo === 'Todos' || d.ticker === vAtivo));
            let tTotal = 0, t12m = 0, mesesUnicos = new Set(), dadosGrafico = {};
            df.forEach(d => { tTotal += d.valor; if(d.is_12m) t12m += d.valor; mesesUnicos.add(d.AnoMes); dadosGrafico[d.AnoMes] = (dadosGrafico[d.AnoMes] || 0) + d.valor; });

            let tMedia = mesesUnicos.size > 0 ? (tTotal / mesesUnicos.size) : 0, labelsBarra = Object.keys(dadosGrafico).sort(), valoresBarra = labelsBarra.map(l => dadosGrafico[l]);
            let tUltimo = labelsBarra.length > 0 ? dadosGrafico[labelsBarra[labelsBarra.length - 1]] : 0;

            document.getElementById('kpi-prov-total').innerText = formataMoedaJS(tTotal); document.getElementById('kpi-prov-12m').innerText = formataMoedaJS(t12m);
            document.getElementById('kpi-prov-media').innerText = formataMoedaJS(tMedia); document.getElementById('kpi-prov-ultimos').innerText = formataMoedaJS(tUltimo);
            
            let tUltimoPct = valInvestidoAtual > 0 ? (tUltimo / valInvestidoAtual) * 100 : 0;
            document.getElementById('kpi-prov-ultimos-pct').innerText = tUltimoPct.toFixed(2).replace('.', ',') + '% sobre investido';

            let layoutProventos = cloneLayout(baseLayout); layoutProventos.showlegend = false;
            Plotly.react('grafico-proventos-mensal', [{ 
                x: labelsBarra, 
                y: valoresBarra, 
                type: 'bar', 
                marker: { color: '#10b981', line: { color: '#059669', width: 1} },
                text: valoresBarra.map(v => formataMoedaJS(v)),
                textposition: 'outside',
                cliponaxis: false
            }], layoutProventos, configPlotly);
        }
        popularFiltros(); atualizarAbaProventos();

        // ================= LÓGICA DE REBALANCEAMENTO (COM TABELA NOVA) =================
        const metricasPlanilha = __DADOS_METRICAS_JS__;
        let segmentTargets = {};
        let segmentsList = [...new Set(dadosCarteira.map(a => a.seguimento || 'Outros'))].sort();
        let defaultPct = segmentsList.length > 0 ? (100 / segmentsList.length) : 0;
        let usaMetricas = Object.keys(metricasPlanilha).length > 0;

        segmentsList.forEach(s => {
            if (usaMetricas && metricasPlanilha[s] !== undefined) {
                segmentTargets[s] = metricasPlanilha[s];
            } else {
                segmentTargets[s] = defaultPct;
            }
        });

        // Variáveis globais para a Tabela e Ordenação
        let rebalanceDataArray = [];
        let curSortCol = 'diff';
        let sortAsc = false; // Começa ordenando decrescente

        function buildRebalanceData() {
            rebalanceDataArray = [];
            let assetsBySeg = {};
            dadosCarteira.forEach(a => {
                let seg = a.seguimento || 'Outros';
                if (!assetsBySeg[seg]) assetsBySeg[seg] = [];
                assetsBySeg[seg].push(a);
            });

            for (let seg in assetsBySeg) {
                let assets = assetsBySeg[seg];
                let targetSegValue = carteiraTotal * (segmentTargets[seg] / 100);
                let targetAssetValue = targetSegValue / assets.length; // Divide igualitário entre ativos do mesmo seguimento

                assets.forEach(a => {
                    let diff = targetAssetValue - a.valor;
                    let pct_atual = carteiraTotal > 0 ? (a.valor / carteiraTotal) * 100 : 0;
                    let target_pct = carteiraTotal > 0 ? (targetAssetValue / carteiraTotal) * 100 : 0;
                    
                    rebalanceDataArray.push({ 
                        ticker: a.ticker || 'N/D', 
                        tipo: a.tipo || 'N/D',
                        seg: seg, 
                        gestora: a.gestora || 'N/D',
                        qtd: a.Qtd || 0,
                        pm: a['Preço_Médio'] || 0,
                        preco: a['Preço_Mercado'] || 0,
                        var: a['Variação'] || 0,
                        current: a.valor || 0,
                        pct_atual: pct_atual,
                        target_pct: target_pct,
                        target: targetAssetValue, 
                        diff: diff
                    });
                });
            }
        }

        function renderTable() {
            // Ordenação
            rebalanceDataArray.sort((a, b) => {
                let valA = a[curSortCol];
                let valB = b[curSortCol];
                if (typeof valA === 'string') valA = valA.toLowerCase();
                if (typeof valB === 'string') valB = valB.toLowerCase();
                if (valA < valB) return sortAsc ? -1 : 1;
                if (valA > valB) return sortAsc ? 1 : -1;
                return 0;
            });

            const tbody = document.getElementById('tabela-rebalanceamento');
            tbody.innerHTML = '';

            rebalanceDataArray.forEach(d => {
                let tr = document.createElement('tr');
                
                let varPct = d.var * 100;
                let varColor = varPct >= 0 ? "var(--green)" : "var(--red)";
                let varText = (varPct >= 0 ? "+" : "") + varPct.toFixed(2).replace(".", ",") + "%";

                let diffText = "";
                let diffColor = "";

                if (varPct >= 20) {
                    diffText = `Vender (Lucro de ${varPct.toFixed(1).replace('.',',')}%)`;
                    diffColor = "var(--red)";
                } 
                else if (d.current <= (d.target * 0.90) && d.preco <= d.pm) {
                    diffText = `Comprar ${formataMoedaJS(d.target - d.current)}`;
                    diffColor = "var(--green)";
                } 
                else if (d.current >= (d.target * 1.10) && d.preco >= d.pm) {
                    diffText = `Vender (Acima do alvo)`;
                    diffColor = "var(--red)";
                } 
                else {
                    diffText = `Falta ${formataMoedaJS(d.target - d.current)}`;
                    diffColor = "var(--text-muted)";
                }

                tr.innerHTML = `
                    <td style="font-weight: 700; color: var(--text-main);">${d.ticker}</td>
                    <td style="color: var(--text-muted);">${d.tipo}</td>
                    <td style="color: var(--text-muted);">${d.seg}</td>
                    <td style="color: var(--text-muted);">${d.gestora}</td>
                    <td style="font-weight: 600; color: var(--primary);">${d.qtd}</td>
                    <td style="font-weight: 500;">${formataMoedaJS(d.pm)}</td>
                    <td style="font-weight: 500;">${formataMoedaJS(d.preco)}</td>
                    <td style="color: ${varColor}; font-weight: 700;">${varText}</td>
                    <td style="font-weight: 600;">${formataMoedaJS(d.current)}</td>
                    <td style="font-weight: 500;">${d.pct_atual.toFixed(2).replace('.',',')}%</td>
                    <td style="color: var(--text-muted);">${d.target_pct.toFixed(2).replace('.',',')}%</td>
                    <td style="color: var(--text-muted);">${formataMoedaJS(d.target)}</td>
                    <td style="color: ${diffColor}; font-weight: 600;">${diffText}</td>
                `;
                tbody.appendChild(tr);
            });
        }

        function sortTable(colName) {
            if (curSortCol === colName) {
                sortAsc = !sortAsc;
            } else {
                curSortCol = colName;
                sortAsc = false;
            }
            renderTable();
        }

        function drawRebalanceTab() {
            const container = document.getElementById('segment-inputs');
            if(container.innerHTML === '') {
                segmentsList.forEach(s => {
                    let div = document.createElement('div');
                    div.innerHTML = `<label style="font-size: 11.5px; color: var(--text-muted); display: block; margin-bottom: 5px; font-weight: 500;">${s}</label>
                                     <input type="number" step="0.5" value="${segmentTargets[s].toFixed(1)}"
                                            onchange="updateTarget('${s}', this.value)"
                                            style="background: var(--bg-color); border: 1px solid var(--border-color); color: var(--text-main); padding: 7px; width: 65px; border-radius: 6px; outline: none; font-weight: 600;"> <span style="color: var(--text-muted); font-size: 13px; font-weight: 500;">%</span>`;
                    container.appendChild(div);
                });
            }

            let totalPct = 0;
            segmentsList.forEach(s => totalPct += segmentTargets[s]);
            const alerta = document.getElementById('soma-alerta');
            alerta.innerText = `Soma atual da Estratégia: ${totalPct.toFixed(1)}%`;
            alerta.style.color = Math.abs(totalPct - 100) < 0.1 ? 'var(--green)' : 'var(--red)';

            buildRebalanceData();
            renderTable();
        }

        function updateTarget(seg, val) {
            segmentTargets[seg] = parseFloat(val) || 0;
            drawRebalanceTab();
        }

        drawRebalanceTab();
        window.addEventListener('resize', () => { window.dispatchEvent(new Event('resize')); });
    </script>
</body>
</html>"""

html_final = html_template.replace("__PATRIMONIO__", str_patrimonio).replace("__PATRIMONIO_ATUAL_NUM__", str(patrimonio_atual))
html_final = html_final.replace("__INVESTIDO__", str_investido).replace("__INVESTIDO_ATUAL_NUM__", str(investido_atual))
html_final = html_final.replace("__LUCRO__", str_lucro).replace("__GANHO_CAPITAL__", str_ganho_capital)
html_final = html_final.replace("__DIVIDENDOS_TOTAL__", str_dividendos_total).replace("__DIVIDENDOS_12M__", str_dividendos_12m)
html_final = html_final.replace("__ULTIMOS_PROVENTOS__", str_ultimos_proventos).replace("__PCT_ULTIMOS_PROVENTOS__", str_pct_ultimos_proventos)
html_final = html_final.replace("__RENTABILIDADE_TOTAL__", str_rentabilidade_total).replace("__RENTABILIDADE_12M__", str_rentabilidade_12m)
html_final = html_final.replace("__COR_LUCRO__", cor_lucro).replace("__COR_GANHO__", cor_ganho)
html_final = html_final.replace("__COR_RENT_12M__", cor_rent_12m).replace("__COR_RENT_TOTAL__", cor_rent_total)
html_final = html_final.replace("__SINAL_RENT_12M__", sinal_rent_12m).replace("__SINAL_RENT_TOTAL__", sinal_rent_total)
html_final = html_final.replace("__COR_VAR_PAT__", cor_var_pat).replace("__VAR_PATRIMONIO__", str_var_patrimonio)
html_final = html_final.replace("__MESES_JS__", meses_js).replace("__VALOR_APLICADO_JS__", valor_aplicado_js).replace("__VALOR_MERCADO_JS__", valor_mercado_js)
html_final = html_final.replace("__DADOS_CARTEIRA_JS__", dados_carteira_js).replace("__DADOS_PROVENTOS_RAW_JS__", dados_proventos_raw_js)
html_final = html_final.replace("__DADOS_METRICAS_JS__", dados_metricas_js)
html_final = html_final.replace("__GANHO_CAPITAL_NUM__", str(ganho_capital))
html_final = html_final.replace("__DIVIDENDOS_TOTAIS_NUM__", str(dividendos_totais))
html_final = html_final.replace("__INSIGHTS_HTML__", html_insights)

nome_arquivo = "index.html"
with open(nome_arquivo, "w", encoding="utf-8") as f: f.write(html_final)
print("Execução Concluída! Planilha Atualizada e Dashboard Premium index.html gerado com sucesso.")
