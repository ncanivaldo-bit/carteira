import os
import json
import time
import gspread
import pandas as pd
import yfinance as yf
from datetime import datetime
from zoneinfo import ZoneInfo
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
    aba_carteira.clear()
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
# 9. GERAÇÃO DO DASHBOARD E INSIGHTS
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
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard de Investimentos</title>
<script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root{
  --bg:#070b16; --bg2:#0b1224; --card:rgba(20,29,50,.72); --card-solid:#141d32;
  --text:#f1f5f9; --muted:#8b9ab5; --green:#10b981; --red:#f43f5e; --amber:#f59e0b;
  --primary:#3b82f6; --violet:#8b5cf6; --cyan:#22d3ee; --border:rgba(255,255,255,.07);
}
*{box-sizing:border-box}
html,body{margin:0}
body{
  background:radial-gradient(1200px 600px at 10% -10%,rgba(59,130,246,.18),transparent 60%),
             radial-gradient(900px 500px at 100% 0%,rgba(139,92,246,.14),transparent 60%),var(--bg);
  background-attachment:fixed;
  color:var(--text);font-family:'Inter','Segoe UI',Roboto,Arial,sans-serif;
  height:100vh;overflow:hidden;display:flex;flex-direction:column;padding:14px 22px;
}
.text-green{color:var(--green)}.text-red{color:var(--red)}
header{display:flex;justify-content:space-between;align-items:center;gap:16px;flex-wrap:wrap;
  padding:6px 0 14px;margin-bottom:14px;border-bottom:1px solid var(--border);flex-shrink:0}
.brand{display:flex;align-items:center;gap:12px}
.logo{width:38px;height:38px;border-radius:11px;background:linear-gradient(135deg,var(--primary),var(--violet));
  display:flex;align-items:center;justify-content:center;font-size:17px;box-shadow:0 6px 20px rgba(59,130,246,.35)}
.brand h1{font-size:16px;margin:0;font-weight:700;letter-spacing:.2px}
.brand small{color:var(--muted);font-size:11.5px}
.nav-links{display:flex;gap:6px;background:rgba(255,255,255,.04);padding:4px;border-radius:12px;border:1px solid var(--border)}
.nav-links a{color:var(--muted);text-decoration:none;font-size:13px;font-weight:600;cursor:pointer;padding:8px 16px;border-radius:9px;transition:.2s;display:flex;gap:7px;align-items:center;white-space:nowrap}
.nav-links a:hover{color:var(--text)}
.nav-links a.active{background:linear-gradient(135deg,var(--primary),#2563eb);color:#fff;box-shadow:0 4px 14px rgba(59,130,246,.4)}
.header-buttons{display:flex;gap:8px;align-items:center}
.btn{background:rgba(255,255,255,.06);color:var(--text);border:1px solid var(--border);padding:9px 14px;border-radius:10px;cursor:pointer;font-size:12.5px;font-weight:600;transition:.2s;font-family:inherit}
.btn:hover{background:rgba(255,255,255,.12)}
.btn-primary{background:linear-gradient(135deg,var(--primary),#2563eb);border:none;color:#fff;box-shadow:0 4px 14px rgba(59,130,246,.35)}
.btn-primary:hover{filter:brightness(1.12);background:linear-gradient(135deg,var(--primary),#2563eb)}
.stamp{font-size:11px;color:var(--muted);text-align:right;line-height:1.3}
.tab-content{display:none;flex-grow:1;flex-direction:column;overflow-y:auto;padding-right:6px;animation:fade .35s ease}
.tab-content.active{display:flex}
@keyframes fade{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}
.tab-content::-webkit-scrollbar,.table-container::-webkit-scrollbar{width:8px;height:8px}
.tab-content::-webkit-scrollbar-thumb,.table-container::-webkit-scrollbar-thumb{background:#26344f;border-radius:8px}
.grid4{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:14px;flex-shrink:0}
.grid-2-1{display:grid;grid-template-columns:2fr 1fr;gap:14px;margin-bottom:14px;flex-shrink:0}
.grid-1-1{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px;flex-shrink:0}
.grid-3{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:14px;flex-shrink:0}
.card{background:var(--card);backdrop-filter:blur(12px);border:1px solid var(--border);border-radius:16px;padding:16px 20px;
  display:flex;flex-direction:column;justify-content:space-between;box-shadow:0 10px 30px -12px rgba(0,0,0,.6);position:relative;overflow:hidden;min-width:0}
.card.hover{transition:.2s}
.card.hover:hover{border-color:rgba(59,130,246,.4);transform:translateY(-2px)}
.kpi-title{font-size:12px;color:var(--muted);font-weight:600;margin-bottom:8px;text-transform:uppercase;letter-spacing:.6px}
.kpi-value{font-size:24px;font-weight:800;display:flex;align-items:baseline;flex-wrap:wrap;gap:4px;letter-spacing:-.3px}
.kpi-sub{font-size:12px;color:var(--muted);display:flex;justify-content:space-between;margin-top:12px;gap:8px}
.chart-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;gap:10px;flex-wrap:wrap}
.chart-title{font-size:14.5px;font-weight:700;display:flex;align-items:center;gap:8px}
.chart-title i{color:var(--primary);font-size:13px}
.chart-sub{font-size:11.5px;color:var(--muted);font-weight:500}
.chart-container{width:100%;height:340px}
.h-md{height:300px}.h-lg{height:420px}.h-sm{height:230px}
.filtros-prov{display:flex;gap:8px;flex-wrap:wrap}
.select-filter,.input-s{background:rgba(255,255,255,.05);border:1px solid var(--border);color:var(--text);padding:7px 12px;border-radius:9px;font-size:12px;outline:none;font-family:inherit}
.select-filter option{background:#141d32}
.input-s:focus,.select-filter:focus{border-color:var(--primary)}
.seg-toggle{display:flex;background:rgba(255,255,255,.05);border-radius:9px;padding:3px;border:1px solid var(--border)}
.seg-toggle button{background:transparent;border:none;color:var(--muted);padding:5px 12px;border-radius:7px;font-size:11.5px;font-weight:600;cursor:pointer;font-family:inherit}
.seg-toggle button.on{background:var(--primary);color:#fff}
.insights{background:linear-gradient(120deg,rgba(59,130,246,.16),rgba(139,92,246,.14));border-left:4px solid var(--primary);margin-bottom:14px;flex-shrink:0}
.insights ul{margin:6px 0 0;padding-left:18px;color:#b4c0d6;font-size:13.5px;line-height:1.55}
.mini-stats{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:6px}
.mini{background:rgba(255,255,255,.04);border:1px solid var(--border);border-radius:12px;padding:10px 12px}
.mini span{display:block;font-size:10.5px;color:var(--muted);font-weight:600;text-transform:uppercase;letter-spacing:.5px}
.mini b{font-size:18px;font-weight:800}
.bar-track{height:6px;background:rgba(255,255,255,.08);border-radius:6px;overflow:hidden;min-width:70px}
.bar-fill{height:100%;border-radius:6px}
.table-container{overflow:auto;max-width:100%;margin-top:10px;flex-grow:1;max-height:560px;border-radius:12px;border:1px solid var(--border)}
table{width:100%;border-collapse:separate;border-spacing:0;text-align:left;font-size:12.8px;min-width:1500px}
th,td{padding:11px 12px;border-bottom:1px solid var(--border);white-space:nowrap}
th{color:var(--muted);font-weight:700;position:sticky;top:0;background:#101828;z-index:2;cursor:pointer;user-select:none;font-size:11.5px;text-transform:uppercase;letter-spacing:.4px}
th:hover{color:var(--text);background:#18233b}
th i{margin-left:5px;font-size:10px;opacity:.6}
td:first-child,th:first-child{position:sticky;left:0;z-index:3;background:#101828;border-right:1px solid var(--border)}
th:first-child{z-index:4}
tbody tr:hover td{background:#18233b}
.badge{display:inline-flex;align-items:center;gap:6px;padding:4px 10px;border-radius:20px;font-size:11.5px;font-weight:700}
.b-buy{background:rgba(16,185,129,.15);color:var(--green)}
.b-sell{background:rgba(244,63,94,.15);color:var(--red)}
.b-wait{background:rgba(148,163,184,.14);color:var(--muted)}
.aporte-row{display:flex;justify-content:space-between;align-items:center;padding:9px 0;border-bottom:1px dashed var(--border);font-size:13px;gap:10px}
.aporte-row:last-child{border:none}
body.privacy .priv{filter:blur(9px);user-select:none}
@media(max-width:1100px){
  body{height:auto;overflow:auto}
  .grid4{grid-template-columns:repeat(2,1fr)}
  .grid-2-1,.grid-1-1,.grid-3{grid-template-columns:1fr}
  .tab-content{overflow:visible}
}
@media(max-width:640px){
  body{padding:10px 12px}
  header{flex-direction:column;align-items:stretch}
  .nav-links{overflow-x:auto}
  .grid4{grid-template-columns:1fr}
  .header-buttons{justify-content:space-between}
  .chart-container{height:300px}
}
</style>
</head>
<body>
<header>
  <div class="brand">
    <div class="logo"><i class="fa-solid fa-chart-line"></i></div>
    <div><h1>Minha Carteira</h1><small>Painel de decisão de investimentos</small></div>
  </div>
  <div class="nav-links">
    <a onclick="mudarAba('resumo')" id="link-resumo" class="active"><i class="fa-solid fa-gauge-high"></i> Resumo</a>
    <a onclick="mudarAba('proventos')" id="link-proventos"><i class="fa-solid fa-coins"></i> Proventos</a>
    <a onclick="mudarAba('gestoras')" id="link-gestoras"><i class="fa-solid fa-building-columns"></i> Gestoras</a>
    <a onclick="mudarAba('rebalanceamento')" id="link-rebalanceamento"><i class="fa-solid fa-scale-balanced"></i> Rebalanceamento</a>
  </div>
  <div class="header-buttons">
    <div class="stamp">Atualizado em<br><b style="color:var(--text)">__ATUALIZADO__</b></div>
    <button class="btn" onclick="togglePrivacy()" id="btn-priv" title="Ocultar valores"><i class="fa-solid fa-eye"></i></button>
    <button class="btn btn-primary" onclick="dispararAtualizacao(this)"><i class="fa-solid fa-rotate"></i> Atualizar Dados</button>
  </div>
</header>

<!-- ================= RESUMO ================= -->
<div id="tab-resumo" class="tab-content active">
  <div class="card insights">
    <div class="kpi-title" style="color:var(--text);font-size:13px;margin:0"><i class="fa-solid fa-wand-magic-sparkles" style="color:var(--primary);margin-right:6px"></i>Insights da Carteira</div>
    <ul id="ai-insights">
      <li id="insight-reco" style="margin-bottom:6px">Analisando oportunidades na carteira...</li>
      __INSIGHTS_HTML__
    </ul>
  </div>

  <div class="grid4">
    <div class="card hover">
      <div class="kpi-title">💰 Patrimônio total</div>
      <div class="kpi-value priv" style="color:var(--primary)">__PATRIMONIO__ <span class="__COR_VAR_PAT__" style="font-size:13.5px;font-weight:700">(__VAR_PATRIMONIO__)</span></div>
      <div id="spark-pat" style="height:44px;margin:6px -8px -4px"></div>
      <div class="kpi-sub" style="margin-top:6px"><span>Valor investido</span><span class="priv" style="font-weight:700;color:var(--text)">__INVESTIDO__</span></div>
    </div>
    <div class="card hover">
      <div class="kpi-title">💲 Lucro total</div>
      <div class="kpi-value priv __COR_LUCRO__">__LUCRO__</div>
      <div class="kpi-sub">
        <div style="display:flex;flex-direction:column;gap:2px"><span>Ganho de capital</span><span class="priv __COR_GANHO__" style="font-size:13px;font-weight:700">__GANHO_CAPITAL__</span></div>
        <div style="display:flex;flex-direction:column;gap:2px;text-align:right"><span>Proventos</span><span class="priv" style="color:var(--text);font-size:13px;font-weight:700">__DIVIDENDOS_TOTAL__</span></div>
      </div>
    </div>
    <div class="card hover">
      <div class="kpi-title">🪙 Últimos proventos</div>
      <div class="kpi-value priv" style="color:var(--cyan)">__ULTIMOS_PROVENTOS__ <span style="font-size:13.5px;font-weight:700;color:var(--muted)">(__PCT_ULTIMOS_PROVENTOS__)</span></div>
      <div class="kpi-sub">
        <div style="display:flex;flex-direction:column;gap:2px"><span>Acumulado</span><span class="priv" style="color:var(--text);font-size:13px;font-weight:700">__DIVIDENDOS_TOTAL__</span></div>
        <div style="display:flex;flex-direction:column;gap:2px;text-align:right"><span>Últimos 12M</span><span class="priv" style="color:var(--text);font-size:13px;font-weight:700">__DIVIDENDOS_12M__</span></div>
      </div>
    </div>
    <div class="card hover">
      <div class="kpi-title">📊 Rentabilidade</div>
      <div style="display:flex;justify-content:space-between;align-items:center">
        <div><div style="font-size:11px;color:var(--muted);margin-bottom:2px">12 meses</div>
          <div class="kpi-value __COR_RENT_12M__" style="font-size:22px">__RENTABILIDADE_12M__ <span style="font-size:13px">__SINAL_RENT_12M__</span></div></div>
        <div style="border-left:1px solid var(--border);padding-left:16px"><div style="font-size:11px;color:var(--muted);margin-bottom:2px">Total</div>
          <div class="__COR_RENT_TOTAL__" style="font-size:18px;font-weight:800">__RENTABILIDADE_TOTAL__ __SINAL_RENT_TOTAL__</div></div>
      </div>
      <div class="kpi-sub"><span>Yield on Cost (12M)</span><span id="kpi-yoc12" style="color:var(--text);font-weight:700">-</span></div>
    </div>
  </div>

  <div class="grid-2-1">
    <div class="card">
      <div class="chart-header">
        <div><div class="chart-title"><i class="fa-solid fa-chart-area"></i> Evolução do Patrimônio</div><div class="chart-sub">Investido vs. valor de mercado</div></div>
        <div class="seg-toggle"><button id="tg-val" class="on" onclick="setEvo('val')">R$</button><button id="tg-rent" onclick="setEvo('rent')">Rentab. %</button></div>
      </div>
      <div id="grafico-linha" class="chart-container priv"></div>
    </div>
    <div class="card">
      <div class="chart-header"><div><div class="chart-title"><i class="fa-solid fa-sun"></i> Composição</div><div class="chart-sub">Tipo › Seguimento › Ativo</div></div></div>
      <div id="grafico-sunburst" class="chart-container"></div>
    </div>
  </div>

  <div class="grid-3">
    <div class="card">
      <div class="chart-header"><div><div class="chart-title"><i class="fa-solid fa-ranking-star"></i> Lucro / Prejuízo por ativo</div><div class="chart-sub">Em R$ sobre o preço médio</div></div></div>
      <div id="grafico-movers" class="chart-container priv"></div>
    </div>
    <div class="card">
      <div class="chart-header"><div><div class="chart-title"><i class="fa-solid fa-chart-pie"></i> Alocação por Tipo</div><div class="chart-sub">Distribuição do patrimônio</div></div></div>
      <div id="grafico-tipo" class="chart-container"></div>
    </div>
    <div class="card">
      <div class="chart-header"><div><div class="chart-title"><i class="fa-solid fa-heart-pulse"></i> Saúde da Carteira</div><div class="chart-sub">Diversificação e concentração</div></div></div>
      <div id="gauge-div" style="height:170px"></div>
      <div class="mini-stats" id="mini-stats"></div>
    </div>
  </div>
</div>

<!-- ================= PROVENTOS ================= -->
<div id="tab-proventos" class="tab-content">
  <div class="grid4">
    <div class="card"><div class="kpi-title">💰 Último mês</div><div class="kpi-value text-green priv" id="kpi-prov-ultimos">R$ 0,00</div><div class="kpi-sub"><span id="kpi-prov-ultimos-pct">0,00% sobre investido</span></div></div>
    <div class="card"><div class="kpi-title">💵 Proventos filtrados</div><div class="kpi-value text-green priv" id="kpi-prov-total">R$ 0,00</div><div class="kpi-sub"><span>Acumulado histórico</span></div></div>
    <div class="card"><div class="kpi-title">📅 Últimos 12 meses</div><div class="kpi-value text-green priv" id="kpi-prov-12m">R$ 0,00</div><div class="kpi-sub"><span id="kpi-prov-yoc">Yield on Cost 12M: -</span></div></div>
    <div class="card"><div class="kpi-title">📈 Média mensal</div><div class="kpi-value priv" style="color:var(--primary)" id="kpi-prov-media">R$ 0,00</div><div class="kpi-sub"><span id="kpi-prov-meses">Por mês com pagamento</span></div></div>
  </div>
  <div class="card" style="margin-bottom:14px">
    <div class="chart-header">
      <div><div class="chart-title"><i class="fa-solid fa-money-bill-trend-up"></i> Evolução Mensal de Proventos</div><div class="chart-sub">Barras: mensal · Linha: acumulado</div></div>
      <div class="filtros-prov">
        <select id="f-tipo" class="select-filter" onchange="atualizarAbaProventos()"></select>
        <select id="f-seguimento" class="select-filter" onchange="atualizarAbaProventos()"></select>
        <select id="f-ativo" class="select-filter" onchange="atualizarAbaProventos()"></select>
      </div>
    </div>
    <div id="grafico-proventos-mensal" class="chart-container h-lg priv"></div>
  </div>
  <div class="grid-1-1">
    <div class="card"><div class="chart-header"><div class="chart-title"><i class="fa-solid fa-calendar-days"></i> Calendário de Proventos</div></div><div id="grafico-heat" class="chart-container h-md priv"></div></div>
    <div class="card"><div class="chart-header"><div class="chart-title"><i class="fa-solid fa-trophy"></i> Quem mais paga</div></div><div id="grafico-prov-ativo" class="chart-container h-md"></div></div>
  </div>
</div>

<!-- ================= GESTORAS ================= -->
<div id="tab-gestoras" class="tab-content">
  <div class="grid-2-1">
    <div class="card"><div class="chart-header"><div><div class="chart-title"><i class="fa-solid fa-table-cells"></i> Alocação por Gestora</div><div class="chart-sub">Tamanho = valor investido atual</div></div></div>
      <div id="grafico-treemap" class="chart-container h-lg"></div></div>
    <div class="card"><div class="chart-header"><div><div class="chart-title"><i class="fa-solid fa-bars-progress"></i> Ranking de concentração</div><div class="chart-sub">% do patrimônio por gestora</div></div></div>
      <div id="grafico-gestora-rank" class="chart-container h-lg"></div></div>
  </div>
</div>

<!-- ================= REBALANCEAMENTO ================= -->
<div id="tab-rebalanceamento" class="tab-content">
  <div class="grid-2-1">
    <div class="card">
      <div class="chart-title"><i class="fa-solid fa-sliders"></i> Estratégia de Alocação por Seguimento</div>
      <div class="chart-sub" style="margin-top:4px">Percentuais importados da aba Metricas. Ajuste para simular cenários.</div>
      <div id="segment-inputs" style="display:flex;gap:14px;flex-wrap:wrap;margin-top:14px"></div>
      <div id="soma-alerta" style="font-size:13px;font-weight:700;margin-top:14px">Soma atual: 100%</div>
    </div>
    <div class="card">
      <div class="chart-title"><i class="fa-solid fa-hand-holding-dollar"></i> Simulador de Aporte</div>
      <div class="chart-sub" style="margin:4px 0 10px">Distribui o valor entre os ativos marcados como "Comprar".</div>
      <div style="display:flex;gap:8px;align-items:center">
        <span style="color:var(--muted);font-weight:700">R$</span>
        <input id="aporte-valor" class="input-s" type="number" min="0" step="100" value="1000" oninput="simularAporte()" style="flex:1">
      </div>
      <div id="aporte-resultado" style="margin-top:10px;max-height:190px;overflow:auto"></div>
    </div>
  </div>
  <div class="card" style="margin-bottom:14px">
    <div class="chart-header"><div><div class="chart-title"><i class="fa-solid fa-chart-column"></i> Atual × Alvo por Seguimento</div><div class="chart-sub">Onde a carteira está desbalanceada</div></div></div>
    <div id="grafico-seg-alvo" class="chart-container h-md"></div>
  </div>
  <div class="card" style="margin-bottom:14px">
    <div class="chart-header">
      <div><div class="chart-title"><i class="fa-solid fa-crosshairs"></i> Radar de Aportes</div><div class="chart-sub">Clique nos títulos para ordenar</div></div>
      <div class="filtros-prov">
        <input id="f-busca" class="input-s" placeholder="Buscar ativo..." oninput="renderTable()">
        <select id="f-status" class="select-filter" onchange="renderTable()">
          <option value="Todos">Todos os Status</option>
          <option value="Comprar">Recomendação: Comprar</option>
          <option value="Vender">Recomendação: Vender</option>
          <option value="Aguardar">Recomendação: Aguardar</option>
        </select>
      </div>
    </div>
    <div class="table-container">
      <table>
        <thead><tr>
          <th onclick="sortTable('ticker')">Ativo <i class="fa-solid fa-sort"></i></th>
          <th onclick="sortTable('actionType')">Ação <i class="fa-solid fa-sort"></i></th>
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
          <th onclick="sortTable('prog')">Progresso do Alvo <i class="fa-solid fa-sort"></i></th>
          <th onclick="sortTable('target')">Valor Alvo <i class="fa-solid fa-sort"></i></th>
          <th onclick="sortTable('diff')">Falta/Sobra (R$) <i class="fa-solid fa-sort"></i></th>
        </tr></thead>
        <tbody id="tabela-rebalanceamento"></tbody>
      </table>
    </div>
  </div>
</div>

<script>
const configPlotly = { displayModeBar:false, responsive:true };
const PAL = ['#3b82f6','#10b981','#8b5cf6','#f59e0b','#ec4899','#22d3ee','#f97316','#14b8a6','#a3e635','#f43f5e'];
const baseLayout = {
  paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
  font:{ color:'#8b9ab5', size:11.5, family:'Inter, Segoe UI' },
  margin:{ t:20, l:55, r:15, b:35 }, hovermode:'x unified',
  hoverlabel:{ bgcolor:'#0b1224', bordercolor:'#26344f', font:{ color:'#f1f5f9' } },
  xaxis:{ showgrid:false, zeroline:false, fixedrange:true },
  yaxis:{ gridcolor:'rgba(255,255,255,.05)', zeroline:false, fixedrange:true },
  colorway:PAL, autosize:true, dragmode:false
};
const cloneLayout = (o) => JSON.parse(JSON.stringify(o));
const formataMoedaJS = (v) => (v||0).toLocaleString('pt-BR',{style:'currency',currency:'BRL'});
const fp = (v,d=2) => (v||0).toFixed(d).replace('.',',') + '%';
const fk = (v) => { const a=Math.abs(v); const s=v<0?'-':''; if(a>=1e6) return s+'R$ '+(a/1e6).toFixed(2).replace('.',',')+' mi'; if(a>=1e3) return s+'R$ '+(a/1e3).toFixed(1).replace('.',',')+' mil'; return s+'R$ '+a.toFixed(0); };
const valInvestidoAtual = __INVESTIDO_ATUAL_NUM__;
const valPatrimonioAtual = __PATRIMONIO_ATUAL_NUM__;

function mudarAba(nomeAba){
  document.querySelectorAll('.tab-content').forEach(el=>el.classList.remove('active'));
  document.querySelectorAll('.nav-links a').forEach(el=>el.classList.remove('active'));
  document.getElementById('tab-'+nomeAba).classList.add('active');
  document.getElementById('link-'+nomeAba).classList.add('active');
  setTimeout(()=>{ document.querySelectorAll('#tab-'+nomeAba+' .chart-container, #tab-'+nomeAba+' #gauge-div').forEach(el=>{ try{Plotly.Plots.resize(el);}catch(e){} }); },60);
}
function togglePrivacy(){
  document.body.classList.toggle('privacy');
  const on = document.body.classList.contains('privacy');
  document.getElementById('btn-priv').innerHTML = on ? '<i class="fa-solid fa-eye-slash"></i>' : '<i class="fa-solid fa-eye"></i>';
}

// ================= DADOS =================
const MESES = __MESES_JS__;
const V_APLICADO = __VALOR_APLICADO_JS__;
const V_MERCADO = __VALOR_MERCADO_JS__;
const dadosCarteira = __DADOS_CARTEIRA_JS__;
let carteiraTotal = 0;
dadosCarteira.forEach(a => { carteiraTotal += a.valor; });

// ================= SPARKLINE =================
Plotly.newPlot('spark-pat',[{ x:MESES, y:V_MERCADO, type:'scatter', mode:'lines', line:{color:'#3b82f6',width:2,shape:'spline'}, fill:'tozeroy', fillcolor:'rgba(59,130,246,.15)', hoverinfo:'skip' }],
 { paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', margin:{t:0,l:0,r:0,b:0}, xaxis:{visible:false,fixedrange:true}, yaxis:{visible:false,fixedrange:true}, dragmode:false },
 { displayModeBar:false, responsive:true, staticPlot:true });

// ================= EVOLUÇÃO =================
function setEvo(modo){
  document.getElementById('tg-val').classList.toggle('on', modo==='val');
  document.getElementById('tg-rent').classList.toggle('on', modo==='rent');
  let lay = cloneLayout(baseLayout); lay.showlegend=true; lay.legend={orientation:'h',y:1.12,x:0};
  if(modo==='val'){
    lay.yaxis.tickformat=',.0f'; lay.yaxis.tickprefix='R$ ';
    Plotly.react('grafico-linha',[
      { x:MESES, y:V_APLICADO, name:'Valor Investido', type:'scatter', mode:'lines', line:{color:'#10b981',width:2.5,dash:'dot',shape:'spline'} },
      { x:MESES, y:V_MERCADO, name:'Valor de Mercado', type:'scatter', mode:'lines', line:{color:'#3b82f6',width:3,shape:'spline'}, fill:'tonexty', fillcolor:'rgba(59,130,246,.14)' }
    ],lay,configPlotly);
  } else {
    const rent = V_MERCADO.map((v,i)=> V_APLICADO[i]>0 ? (v/V_APLICADO[i]-1)*100 : 0);
    lay.yaxis.ticksuffix='%'; lay.yaxis.zeroline=true; lay.yaxis.zerolinecolor='rgba(255,255,255,.25)';
    Plotly.react('grafico-linha',[
      { x:MESES, y:rent, name:'Rentabilidade s/ investido (%)', type:'scatter', mode:'lines', line:{color:'#22d3ee',width:3,shape:'spline'}, fill:'tozeroy', fillcolor:'rgba(34,211,238,.12)', hovertemplate:'%{y:.2f}%' }
    ],lay,configPlotly);
  }
}
setEvo('val');

// ================= SUNBURST =================
let ids=["Carteira"],labels=["Carteira"],parents=[""],values=[carteiraTotal],totaisTipo={},totaisSeg={};
dadosCarteira.forEach(a=>{ let tipo=a.tipo||'Outros',seg=a.seguimento||'Outros',segId=tipo+" - "+seg; totaisTipo[tipo]=(totaisTipo[tipo]||0)+a.valor; if(!totaisSeg[segId]) totaisSeg[segId]={label:seg,parent:tipo,val:0}; totaisSeg[segId].val+=a.valor; });
for(let t in totaisTipo){ ids.push(t); labels.push(t); parents.push("Carteira"); values.push(totaisTipo[t]); }
for(let s in totaisSeg){ ids.push(s); labels.push(totaisSeg[s].label); parents.push(totaisSeg[s].parent); values.push(totaisSeg[s].val); }
dadosCarteira.forEach(a=>{ let tipo=a.tipo||'Outros',seg=a.seguimento||'Outros',segId=tipo+" - "+seg; ids.push(a.ticker); labels.push(a.ticker); parents.push(segId); values.push(a.valor); });
let customText=[],customHover=[];
for(let i=0;i<values.length;i++){
  let strMoeda=formataMoedaJS(values[i]);
  if(i===0){ let sp=formataMoedaJS(valPatrimonioAtual); customText.push(`<b>${labels[i]}</b><br>${sp}`); customHover.push(`<b>${labels[i]}</b><br>Patrimônio Total: ${sp}`); }
  else { let pct=carteiraTotal>0?(values[i]/carteiraTotal)*100:0; let sp=fp(pct); customText.push(`${labels[i]}<br>${sp}`); customHover.push(`<b>${labels[i]}</b><br>Valor: ${strMoeda}<br>Participação: ${sp}`); }
}
let layoutSunburst=cloneLayout(baseLayout); layoutSunburst.margin={t:5,l:5,r:5,b:5}; delete layoutSunburst.xaxis; delete layoutSunburst.yaxis;
Plotly.newPlot('grafico-sunburst',[{type:"sunburst",ids:ids,labels:labels,parents:parents,values:values,branchvalues:"total",text:customText,hovertext:customHover,textinfo:"text",hoverinfo:"text",marker:{line:{width:1.5,color:'#0b1224'}}}],layoutSunburst,configPlotly);

// ================= MOVERS (LUCRO R$ POR ATIVO) =================
(function(){
  let arr = dadosCarteira.map(a=>({ t:a.ticker, l:(a.valor||0) - (a.Qtd||0)*(a['Preço_Médio']||0) })).sort((x,y)=>y.l-x.l);
  let sel = arr.length>14 ? arr.slice(0,7).concat(arr.slice(-7)) : arr;
  sel = sel.slice().reverse();
  let lay = cloneLayout(baseLayout); lay.margin={t:5,l:62,r:20,b:30}; lay.hovermode='closest';
  lay.xaxis={gridcolor:'rgba(255,255,255,.05)',zeroline:true,zerolinecolor:'rgba(255,255,255,.3)',fixedrange:true,tickformat:',.0f'};
  lay.yaxis={automargin:true,fixedrange:true,showgrid:false};
  Plotly.newPlot('grafico-movers',[{ type:'bar', orientation:'h', x:sel.map(s=>s.l), y:sel.map(s=>s.t),
    marker:{ color:sel.map(s=>s.l>=0?'#10b981':'#f43f5e'), line:{width:0} },
    hovertemplate:'<b>%{y}</b><br>%{x:,.2f} R$<extra></extra>' }],lay,configPlotly);
})();

// ================= ALOCAÇÃO POR TIPO (DONUT) =================
(function(){
  let lay=cloneLayout(baseLayout); lay.margin={t:5,l:5,r:5,b:5}; delete lay.xaxis; delete lay.yaxis; lay.showlegend=true; lay.legend={orientation:'h',y:-0.05,font:{size:11}};
  let tl=Object.keys(totaisTipo), tv=tl.map(k=>totaisTipo[k]);
  Plotly.newPlot('grafico-tipo',[{type:'pie',labels:tl,values:tv,hole:.62,sort:true,textinfo:'percent',textfont:{color:'#fff',size:12},
    marker:{colors:PAL,line:{color:'#0b1224',width:2}},hovertemplate:'<b>%{label}</b><br>%{value:,.2f}<br>%{percent}<extra></extra>'}],lay,configPlotly);
})();

// ================= SAÚDE DA CARTEIRA =================
(function(){
  const n=dadosCarteira.length;
  const pesos=dadosCarteira.map(a=>carteiraTotal>0?a.valor/carteiraTotal:0).sort((a,b)=>b-a);
  const hhi=pesos.reduce((s,w)=>s+w*w,0);
  const nEf=hhi>0?1/hhi:0;
  const top1=(pesos[0]||0)*100, top5=pesos.slice(0,5).reduce((s,w)=>s+w,0)*100;
  const noLucro=dadosCarteira.filter(a=>(a['Variação']||0)>0).length;
  const score=n>0?Math.min(100,(nEf/n)*100):0;
  const cor=score>=75?'#10b981':(score>=50?'#f59e0b':'#f43f5e');
  Plotly.newPlot('gauge-div',[{type:'indicator',mode:'gauge+number',value:score,number:{suffix:'%',font:{size:30,color:'#f1f5f9'}},
    gauge:{axis:{range:[0,100],tickcolor:'#475569',tickfont:{size:9}},bar:{color:cor,thickness:.3},bgcolor:'rgba(255,255,255,.05)',borderwidth:0,
    steps:[{range:[0,50],color:'rgba(244,63,94,.12)'},{range:[50,75],color:'rgba(245,158,11,.12)'},{range:[75,100],color:'rgba(16,185,129,.12)'}]}}],
    {paper_bgcolor:'rgba(0,0,0,0)',margin:{t:20,l:25,r:25,b:0},font:{family:'Inter'}},configPlotly);
  const st=(l,v,c)=>`<div class="mini"><span>${l}</span><b style="color:${c||'var(--text)'}">${v}</b></div>`;
  document.getElementById('mini-stats').innerHTML =
    st('Ativos',n)+st('N° efetivo',nEf.toFixed(1).replace('.',','))+
    st('Maior posição',fp(top1,1), top1>25?'var(--red)':(top1>15?'var(--amber)':'var(--green)'))+
    st('Top 5',fp(top5,1), top5>70?'var(--amber)':'var(--green)')+
    st('No lucro',noLucro+'/'+n,'var(--green)')+st('No prejuízo',(n-noLucro)+'/'+n,(n-noLucro)>0?'var(--red)':'var(--green)');
  const li=(t)=>{ const e=document.createElement('li'); e.style.marginBottom='6px'; e.innerHTML=t; document.getElementById('ai-insights').appendChild(e); };
  if(n>0){
    const mx=dadosCarteira.slice().sort((a,b)=>b.valor-a.valor)[0];
    li(`🎯 Sua maior posição é <b>${mx.ticker}</b> com <b>${fp(top1,1)}</b> do patrimônio${top1>20?' — atenção à concentração.':'.'}`);
    li(`🧩 Diversificação efetiva equivale a <b>${nEf.toFixed(1).replace('.',',')}</b> ativos de peso igual (de ${n} na carteira).`);
    const pior=dadosCarteira.slice().sort((a,b)=>(a['Variação']||0)-(b['Variação']||0))[0];
    if((pior['Variação']||0)<0) li(`⚠️ O ativo mais pressionado é <b>${pior.ticker}</b>, com <b>${fp((pior['Variação']||0)*100)}</b> sobre o preço médio.`);
  }
})();

// ================= GESTORAS =================
let tmIds=["Alocação"],tmLabels=["Alocação"],tmParents=[""],tmValues=[carteiraTotal],totaisGestora={};
dadosCarteira.forEach(a=>{ let g=a.gestora||'Outros'; totaisGestora[g]=(totaisGestora[g]||0)+a.valor; });
for(let g in totaisGestora){ tmIds.push(g); tmLabels.push(g); tmParents.push("Alocação"); tmValues.push(totaisGestora[g]); }
dadosCarteira.forEach(a=>{ let g=a.gestora||'Outros'; tmIds.push(a.ticker+" ("+g+")"); tmLabels.push(a.ticker); tmParents.push(g); tmValues.push(a.valor); });
let tmText=[],tmHover=[];
for(let i=0;i<tmValues.length;i++){
  let strMoeda=formataMoedaJS(tmValues[i]);
  if(i===0){ let sp=formataMoedaJS(valPatrimonioAtual); tmText.push(`<b>${tmLabels[i]}</b><br>${sp}`); tmHover.push(`<b>${tmLabels[i]}</b><br>Patrimônio Total: ${sp}`); }
  else { let pct=carteiraTotal>0?(tmValues[i]/carteiraTotal)*100:0; let sp=fp(pct); tmText.push(`<b>${tmLabels[i]}</b><br>${sp}`); tmHover.push(`<b>${tmLabels[i]}</b><br>Valor: ${strMoeda}<br>Participação: ${sp}`); }
}
let layoutTreemap=cloneLayout(baseLayout); layoutTreemap.margin={t:5,l:5,r:5,b:5}; delete layoutTreemap.xaxis; delete layoutTreemap.yaxis;
Plotly.newPlot('grafico-treemap',[{type:"treemap",ids:tmIds,labels:tmLabels,parents:tmParents,values:tmValues,branchvalues:"total",text:tmText,hovertext:tmHover,textinfo:"text",hoverinfo:"text",marker:{line:{width:1.5,color:'#0b1224'}},pathbar:{visible:false}}],layoutTreemap,configPlotly);
(function(){
  let g=Object.keys(totaisGestora).map(k=>({k:k,v:totaisGestora[k]})).sort((a,b)=>a.v-b.v);
  let lay=cloneLayout(baseLayout); lay.margin={t:5,l:10,r:60,b:30}; lay.hovermode='closest';
  lay.xaxis={visible:false,fixedrange:true}; lay.yaxis={automargin:true,showgrid:false,fixedrange:true};
  Plotly.newPlot('grafico-gestora-rank',[{type:'bar',orientation:'h',x:g.map(d=>d.v),y:g.map(d=>d.k),
    text:g.map(d=>fp(carteiraTotal>0?d.v/carteiraTotal*100:0,1)),textposition:'outside',cliponaxis:false,textfont:{color:'#f1f5f9'},
    marker:{color:g.map((d,i)=>PAL[i%PAL.length])},hovertemplate:'<b>%{y}</b><br>%{x:,.2f}<extra></extra>'}],lay,configPlotly);
})();

// ================= PROVENTOS =================
const dadosProvRaw = __DADOS_PROVENTOS_RAW_JS__;
function popularFiltros(){
  let tipos=new Set(),segs=new Set(),ativos=new Set();
  dadosProvRaw.forEach(d=>{ tipos.add(d.tipo); segs.add(d.seguimento); ativos.add(d.ticker); });
  const addOp=(sel,val,text)=>{ sel.add(new Option(text,val)); };
  addOp(document.getElementById('f-tipo'),'Todos','Tipo: Todos'); Array.from(tipos).sort().forEach(v=>addOp(document.getElementById('f-tipo'),v,v));
  addOp(document.getElementById('f-seguimento'),'Todos','Seguimento: Todos'); Array.from(segs).sort().forEach(v=>addOp(document.getElementById('f-seguimento'),v,v));
  addOp(document.getElementById('f-ativo'),'Todos','Ativo: Todos'); Array.from(ativos).sort().forEach(v=>addOp(document.getElementById('f-ativo'),v,v));
}
function atualizarAbaProventos(){
  const vTipo=document.getElementById('f-tipo').value,vSeg=document.getElementById('f-seguimento').value,vAtivo=document.getElementById('f-ativo').value;
  let df=dadosProvRaw.filter(d=>(vTipo==='Todos'||d.tipo===vTipo)&&(vSeg==='Todos'||d.seguimento===vSeg)&&(vAtivo==='Todos'||d.ticker===vAtivo));
  let tTotal=0,t12m=0,mesesUnicos=new Set(),dadosGrafico={};
  df.forEach(d=>{ tTotal+=d.valor; if(d.is_12m) t12m+=d.valor; mesesUnicos.add(d.AnoMes); dadosGrafico[d.AnoMes]=(dadosGrafico[d.AnoMes]||0)+d.valor; });
  let tMedia=mesesUnicos.size>0?(tTotal/mesesUnicos.size):0, labelsBarra=Object.keys(dadosGrafico).sort(), valoresBarra=labelsBarra.map(l=>dadosGrafico[l]);
  let tUltimo=labelsBarra.length>0?dadosGrafico[labelsBarra[labelsBarra.length-1]]:0;
  document.getElementById('kpi-prov-total').innerText=formataMoedaJS(tTotal);
  document.getElementById('kpi-prov-12m').innerText=formataMoedaJS(t12m);
  document.getElementById('kpi-prov-media').innerText=formataMoedaJS(tMedia);
  document.getElementById('kpi-prov-ultimos').innerText=formataMoedaJS(tUltimo);
  document.getElementById('kpi-prov-meses').innerText='Em '+mesesUnicos.size+' meses com pagamento';
  let tUltimoPct=valInvestidoAtual>0?(tUltimo/valInvestidoAtual)*100:0;
  document.getElementById('kpi-prov-ultimos-pct').innerText=fp(tUltimoPct)+' sobre investido';
  let yoc12=valInvestidoAtual>0?(t12m/valInvestidoAtual)*100:0;
  document.getElementById('kpi-prov-yoc').innerText='Yield on Cost 12M: '+fp(yoc12);
  document.getElementById('kpi-yoc12').innerText=fp(valInvestidoAtual>0?(dadosProvRaw.filter(d=>d.is_12m).reduce((s,d)=>s+d.valor,0)/valInvestidoAtual)*100:0);

  let acc=0, accArr=valoresBarra.map(v=>{ acc+=v; return acc; });
  let lay=cloneLayout(baseLayout); lay.showlegend=false;
  lay.yaxis.tickprefix='R$ '; lay.margin={t:20,l:65,r:70,b:40};
  lay.yaxis2={overlaying:'y',side:'right',showgrid:false,zeroline:false,fixedrange:true,tickprefix:'R$ ',tickfont:{color:'#22d3ee'}};
  Plotly.react('grafico-proventos-mensal',[
    { x:labelsBarra, y:valoresBarra, type:'bar', name:'Mensal', marker:{color:'#10b981',line:{color:'#059669',width:1}},
      text: labelsBarra.length<=18 ? valoresBarra.map(v=>fk(v)) : null, textposition:'outside', cliponaxis:false, textfont:{size:10.5,color:'#cbd5e1'} },
    { x:labelsBarra, y:accArr, type:'scatter', mode:'lines', name:'Acumulado', yaxis:'y2', line:{color:'#22d3ee',width:2.5,shape:'spline'} }
  ],lay,configPlotly);

  const meses=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  let anos=[...new Set(labelsBarra.filter(l=>/^[0-9]{4}-[0-9]{2}$/.test(l)).map(l=>l.slice(0,4)))].sort();
  let z=anos.map(a=>meses.map((m,i)=>{ const k=a+'-'+String(i+1).padStart(2,'0'); return dadosGrafico[k]!==undefined?dadosGrafico[k]:null; }));
  let lh=cloneLayout(baseLayout); lh.hovermode='closest'; lh.margin={t:10,l:45,r:10,b:30};
  lh.xaxis={fixedrange:true,showgrid:false}; lh.yaxis={fixedrange:true,showgrid:false,type:'category',autorange:'reversed'};
  Plotly.react('grafico-heat',[{type:'heatmap',x:meses,y:anos,z:z,xgap:3,ygap:3,hoverongaps:false,
    colorscale:[[0,'rgba(16,185,129,.12)'],[1,'#10b981']],showscale:false,
    hovertemplate:'<b>%{y} · %{x}</b><br>R$ %{z:,.2f}<extra></extra>'}],lh,configPlotly);

  let porAtivo={}; df.forEach(d=>{ porAtivo[d.ticker]=(porAtivo[d.ticker]||0)+d.valor; });
  let pa=Object.keys(porAtivo).map(k=>({k:k,v:porAtivo[k]})).sort((a,b)=>b.v-a.v);
  if(pa.length>9){ const resto=pa.slice(8).reduce((s,d)=>s+d.v,0); pa=pa.slice(0,8); pa.push({k:'Outros',v:resto}); }
  let lp=cloneLayout(baseLayout); lp.margin={t:5,l:5,r:5,b:5}; delete lp.xaxis; delete lp.yaxis; lp.showlegend=true; lp.legend={font:{size:11}};
  Plotly.react('grafico-prov-ativo',[{type:'pie',labels:pa.map(d=>d.k),values:pa.map(d=>d.v),hole:.58,textinfo:'percent',
    marker:{colors:PAL,line:{color:'#0b1224',width:2}},hovertemplate:'<b>%{label}</b><br>R$ %{value:,.2f}<br>%{percent}<extra></extra>'}],lp,configPlotly);
}
popularFiltros(); atualizarAbaProventos();

// ================= REBALANCEAMENTO =================
const metricasPlanilha = __DADOS_METRICAS_JS__;
let segmentTargets={};
let segmentsList=[...new Set(dadosCarteira.map(a=>a.seguimento||'Outros'))].sort();
let defaultPct=segmentsList.length>0?(100/segmentsList.length):0;
let usaMetricas=Object.keys(metricasPlanilha).length>0;
segmentsList.forEach(s=>{ segmentTargets[s]=(usaMetricas&&metricasPlanilha[s]!==undefined)?metricasPlanilha[s]:defaultPct; });
let rebalanceDataArray=[]; let curSortCol='diff'; let sortAsc=false;

function buildRebalanceData(){
  rebalanceDataArray=[]; let assetsBySeg={};
  dadosCarteira.forEach(a=>{ let seg=a.seguimento||'Outros'; if(!assetsBySeg[seg]) assetsBySeg[seg]=[]; assetsBySeg[seg].push(a); });
  for(let seg in assetsBySeg){
    let assets=assetsBySeg[seg];
    let targetSegValue=carteiraTotal*(segmentTargets[seg]/100);
    let targetAssetValue=targetSegValue/assets.length;
    assets.forEach(a=>{
      let pct_atual=carteiraTotal>0?(a.valor/carteiraTotal)*100:0;
      let target_pct=carteiraTotal>0?(targetAssetValue/carteiraTotal)*100:0;
      let varPct=(a['Variação']||0)*100;
      let current=a.valor||0; let preco=a['Preço_Mercado']||0; let pm=a['Preço_Médio']||0;
      let actionType="Aguardar"; let diffValue=targetAssetValue-current;
      if(varPct>=20){ actionType="Vender"; }
      else if(current<=(targetAssetValue*0.90)&&preco<=pm){ actionType="Comprar"; }
      else if(current>=(targetAssetValue*1.10)&&preco>=pm){ actionType="Vender"; }
      rebalanceDataArray.push({ ticker:a.ticker||'N/D', tipo:a.tipo||'N/D', seg:seg, gestora:a.gestora||'N/D', qtd:a.Qtd||0, pm:pm, preco:preco,
        var:a['Variação']||0, current:current, pct_atual:pct_atual, target_pct:target_pct, target:targetAssetValue, diff:diffValue, actionType:actionType,
        prog: targetAssetValue>0 ? (current/targetAssetValue)*100 : 0 });
    });
  }
  let comprarList=rebalanceDataArray.filter(d=>d.actionType==='Comprar').map(d=>d.ticker);
  let venderList=rebalanceDataArray.filter(d=>d.actionType==='Vender').map(d=>d.ticker);
  let htmlReco="";
  if(comprarList.length>0){ htmlReco+=`🟢 <b>Oportunidades de Compra:</b> ${comprarList.join(', ')}. <br>`; }
  else { htmlReco+=`🟢 <b>Oportunidades de Compra:</b> Nenhuma no momento. Ativos no alvo ou esticados. <br>`; }
  if(venderList.length>0){ htmlReco+=`🔴 <b>Analisar Venda (Lucro ou Acima do Alvo):</b> ${venderList.join(', ')}.`; }
  document.getElementById('insight-reco').innerHTML=htmlReco;
}

function renderTable(){
  let filtroStatus=document.getElementById('f-status').value;
  let busca=(document.getElementById('f-busca').value||'').toLowerCase().trim();
  let dadosFiltrados=rebalanceDataArray.filter(d=>{
    if(busca && String(d.ticker).toLowerCase().indexOf(busca)===-1) return false;
    return filtroStatus==='Todos' || d.actionType===filtroStatus;
  });
  dadosFiltrados.sort((a,b)=>{
    let valA=a[curSortCol],valB=b[curSortCol];
    if(typeof valA==='string') valA=valA.toLowerCase();
    if(typeof valB==='string') valB=valB.toLowerCase();
    if(valA<valB) return sortAsc?-1:1;
    if(valA>valB) return sortAsc?1:-1;
    return 0;
  });
  const tbody=document.getElementById('tabela-rebalanceamento'); tbody.innerHTML='';
  dadosFiltrados.forEach(d=>{
    let tr=document.createElement('tr');
    let varPct=d.var*100;
    let varColor=varPct>=0?"var(--green)":"var(--red)";
    let varText=(varPct>=0?"+":"")+fp(varPct);
    let diffText="",diffColor="",badge="";
    if(d.actionType==="Vender"&&varPct>=20){ diffText=`Vender (Lucro de ${varPct.toFixed(1).replace('.',',')}%)`; diffColor="var(--red)"; }
    else if(d.actionType==="Comprar"){ diffText=`Comprar ${formataMoedaJS(d.target-d.current)}`; diffColor="var(--green)"; }
    else if(d.actionType==="Vender"){ diffText=`Vender (Acima do alvo)`; diffColor="var(--red)"; }
    else {
      if(d.diff>0) diffText=`Falta ${formataMoedaJS(d.diff)}`;
      else if(d.diff<0) diffText=`Sobra ${formataMoedaJS(Math.abs(d.diff))}`;
      else diffText="No Alvo";
      diffColor="var(--muted)";
    }
    if(d.actionType==="Comprar") badge='<span class="badge b-buy"><i class="fa-solid fa-arrow-trend-up"></i> Comprar</span>';
    else if(d.actionType==="Vender") badge='<span class="badge b-sell"><i class="fa-solid fa-arrow-trend-down"></i> Vender</span>';
    else badge='<span class="badge b-wait"><i class="fa-solid fa-hourglass-half"></i> Aguardar</span>';
    let pg=Math.min(d.prog,150); let pgColor=d.prog<90?'#3b82f6':(d.prog<=110?'#10b981':'#f59e0b');
    tr.innerHTML=`
      <td style="font-weight:800;color:var(--text)">${d.ticker}</td>
      <td>${badge}</td>
      <td style="color:var(--muted)">${d.tipo}</td>
      <td style="color:var(--muted)">${d.seg}</td>
      <td style="color:var(--muted)">${d.gestora}</td>
      <td style="font-weight:600;color:var(--primary)">${d.qtd}</td>
      <td class="priv">${formataMoedaJS(d.pm)}</td>
      <td class="priv">${formataMoedaJS(d.preco)}</td>
      <td style="color:${varColor};font-weight:700">${varText}</td>
      <td class="priv" style="font-weight:600">${formataMoedaJS(d.current)}</td>
      <td>${fp(d.pct_atual)}</td>
      <td style="color:var(--muted)">${fp(d.target_pct)}</td>
      <td><div style="display:flex;align-items:center;gap:8px"><div class="bar-track" style="width:90px"><div class="bar-fill" style="width:${(pg/150*100).toFixed(0)}%;background:${pgColor}"></div></div><span style="font-size:11.5px;color:var(--muted)">${d.prog.toFixed(0)}%</span></div></td>
      <td class="priv" style="color:var(--muted)">${formataMoedaJS(d.target)}</td>
      <td class="priv" style="color:${diffColor};font-weight:600">${diffText}</td>`;
    tbody.appendChild(tr);
  });
}
function sortTable(colName){
  if(curSortCol===colName){ sortAsc=!sortAsc; } else { curSortCol=colName; sortAsc=false; }
  renderTable();
}

function desenharSegAlvo(){
  let atual={}; dadosCarteira.forEach(a=>{ let s=a.seguimento||'Outros'; atual[s]=(atual[s]||0)+a.valor; });
  let segs=segmentsList;
  let lay=cloneLayout(baseLayout); lay.barmode='group'; lay.showlegend=true; lay.legend={orientation:'h',y:1.15,x:0};
  lay.yaxis.ticksuffix='%'; lay.margin={t:20,l:50,r:15,b:70}; lay.xaxis.tickangle=-25;
  Plotly.react('grafico-seg-alvo',[
    {type:'bar',name:'Atual',x:segs,y:segs.map(s=>carteiraTotal>0?(atual[s]||0)/carteiraTotal*100:0),marker:{color:'#3b82f6'},hovertemplate:'%{y:.2f}%'},
    {type:'bar',name:'Alvo',x:segs,y:segs.map(s=>segmentTargets[s]||0),marker:{color:'#8b5cf6'},hovertemplate:'%{y:.2f}%'}
  ],lay,configPlotly);
}

function simularAporte(){
  const v=parseFloat(document.getElementById('aporte-valor').value)||0;
  const box=document.getElementById('aporte-resultado');
  const alvos=rebalanceDataArray.filter(d=>d.actionType==='Comprar'&&(d.target-d.current)>0);
  if(alvos.length===0){ box.innerHTML='<div class="chart-sub" style="padding:10px 0">Nenhum ativo com sinal de compra no momento.</div>'; return; }
  const falta=alvos.reduce((s,d)=>s+(d.target-d.current),0);
  let html='';
  alvos.slice().sort((a,b)=>(b.target-b.current)-(a.target-a.current)).forEach(d=>{
    const need=d.target-d.current;
    const aloc=Math.min(need, v*need/falta);
    const cotas=d.preco>0?Math.floor(aloc/d.preco):0;
    html+=`<div class="aporte-row"><div><b>${d.ticker}</b><div class="chart-sub">${d.seg}</div></div>
      <div style="text-align:right"><b style="color:var(--green)" class="priv">${formataMoedaJS(aloc)}</b><div class="chart-sub">≈ ${cotas} cotas</div></div></div>`;
  });
  if(v>falta) html+=`<div class="chart-sub" style="padding-top:8px">Sobra <b class="priv">${formataMoedaJS(v-falta)}</b> após atingir todos os alvos.</div>`;
  box.innerHTML=html;
}

function drawRebalanceTab(){
  const container=document.getElementById('segment-inputs');
  if(container.innerHTML===''){
    segmentsList.forEach(s=>{
      let div=document.createElement('div');
      div.innerHTML=`<label style="font-size:11.5px;color:var(--muted);display:block;margin-bottom:5px;font-weight:600">${s}</label>
        <input class="input-s" type="number" step="0.5" value="${segmentTargets[s].toFixed(1)}" onchange="updateTarget('${s}', this.value)" style="width:72px;font-weight:700"> <span style="color:var(--muted);font-size:13px">%</span>`;
      container.appendChild(div);
    });
  }
  let totalPct=0; segmentsList.forEach(s=>totalPct+=segmentTargets[s]);
  const alerta=document.getElementById('soma-alerta');
  alerta.innerText=`Soma atual da Estratégia: ${totalPct.toFixed(1)}% ${Math.abs(totalPct-100)<0.1?'✔':'(deve somar 100%)'}`;
  alerta.style.color=Math.abs(totalPct-100)<0.1?'var(--green)':'var(--red)';
  buildRebalanceData(); renderTable(); desenharSegAlvo(); simularAporte();
}
function updateTarget(seg,val){ segmentTargets[seg]=parseFloat(val)||0; drawRebalanceTab(); }
drawRebalanceTab();

// ================= ATUALIZAR VIA GITHUB API =================
async function dispararAtualizacao(btn){
  const GITHUB_USER='ncanivaldo-bit';
  const GITHUB_REPO='carteira';
  const WORKFLOW_FILE='atualiza_dashboard.yml';
  let token=localStorage.getItem('github_token');
  if(!token){
    token=prompt("🔒 Primeira vez atualizando por aqui!\\n\\nPor favor, insira seu Personal Access Token do GitHub (começa com ghp_):");
    if(!token){ alert("Atualização cancelada. O Token é necessário."); return; }
    localStorage.setItem('github_token',token);
  }
  const textoOriginal=btn.innerHTML;
  btn.innerHTML='<i class="fa-solid fa-spinner fa-spin"></i> Enviando comando...';
  btn.disabled=true;
  try{
    const response=await fetch(`https://api.github.com/repos/${GITHUB_USER}/${GITHUB_REPO}/actions/workflows/${WORKFLOW_FILE}/dispatches`,{
      method:'POST',
      headers:{ 'Accept':'application/vnd.github.v3+json', 'Authorization':`token ${token}` },
      body:JSON.stringify({ ref:'main' })
    });
    if(response.ok||response.status===204){
      btn.innerHTML='<i class="fa-solid fa-check"></i> Atualizando...';
      btn.style.background='var(--green)';
      alert("🚀 Comando enviado com sucesso!\\n\\nO servidor do GitHub já está rodando o script. Recarregue esta página em cerca de 40 a 60 segundos.");
    } else {
      const err=await response.json();
      throw new Error(err.message||'Erro na autenticação');
    }
  }catch(error){
    console.error(error);
    alert(`❌ Erro ao atualizar: ${error.message}\\n\\nSe o seu token for inválido, vamos resetá-lo.`);
    localStorage.removeItem('github_token');
    btn.innerHTML=textoOriginal;
    btn.disabled=false;
  }
}
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
html_final = html_final.replace("__INSIGHTS_HTML__", html_insights)
html_final = html_final.replace("__ATUALIZADO__", datetime.now(ZoneInfo("America/Sao_Paulo")).strftime('%d/%m/%Y %H:%M'))
nome_arquivo = "index.html"
with open(nome_arquivo, "w", encoding="utf-8") as f: f.write(html_final)
print("Execução Concluída! Planilha Atualizada e Dashboard Premium index.html gerado com sucesso.")
