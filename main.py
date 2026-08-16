import os
import json
import time
import gspread
import pandas as pd
import numpy as np
import yfinance as yf
from datetime import datetime
from dateutil.relativedelta import relativedelta
from google.oauth2.service_account import Credentials

# ==============================================================================
# 1. CONFIGURAÇÕES E AUTENTICAÇÃO
# ==============================================================================
print("🚀 Iniciando Motor de Inteligência de Carteira V2...")

CHAVE_JSON_STR = os.environ.get('GCP_CREDENTIALS')
if not CHAVE_JSON_STR:
    raise ValueError("ERRO: Variável 'GCP_CREDENTIALS' não encontrada.")

scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = Credentials.from_service_account_info(json.loads(CHAVE_JSON_STR), scopes=scopes)
gc = gspread.authorize(creds)

planilha = gc.open("base")
aba_mov = planilha.worksheet("Movimentacao")
aba_hist = planilha.worksheet("Hist_Precos")

# Dicionários
conversao_tickers = {"MALL11": "PMLL11", "CVBI11": "PCIP11", "BOML": "BPML11"}
ticker_origem = {v: k for k, v in conversao_tickers.items()}
mes_atual_str = datetime.now().strftime('%Y-%m')

# ==============================================================================
# 2. CARREGAMENTO E HIGIENIZAÇÃO DE DADOS
# ==============================================================================
print("📊 Carregando e limpando dados...")
df_mov = pd.DataFrame(aba_mov.get_all_records(value_render_option='UNFORMATTED_VALUE'))

def parse_date(val):
    if pd.isna(val) or val == '': return pd.NaT
    if isinstance(val, (int, float)): return pd.to_datetime('1899-12-30') + pd.to_timedelta(val, unit='D')
    return pd.to_datetime(val, dayfirst=True, errors='coerce')

df_mov['Data'] = df_mov.get('Data', pd.Series()).apply(parse_date)
df_mov['Ticker'] = df_mov.get('Ticker', '').astype(str).str.strip().str.upper()

# ==============================================================================
# 3. MOTOR DE PREÇO MÉDIO E CONSTRUÇÃO DE PORTFÓLIO
# ==============================================================================
print("⚙️ Calculando Preço Médio e Posições (Algoritmo FIFO/PM)...")
df_mov_calc = df_mov.sort_values(by=['Data', 'Movimentação']).reset_index(drop=True)

carteira = {}
historico_pm = []
proventos = []

for idx, row in df_mov_calc.iterrows():
    tk = row['Ticker']
    if not tk: continue

    mov = str(row.get('Movimentação', '')).strip().title()
    es = str(row.get('Entrada/Saída', row.get('Entrada/Saida', ''))).strip().title()
    qtd = float(row.get('Quantidade', 0))
    valor_op = float(str(row.get('Valor da Operação', 0)).replace('R$', '').replace(',', '.'))
    preco_unit = float(str(row.get('Preço unitário', 0)).replace('R$', '').replace(',', '.'))
    
    if valor_op == 0 and preco_unit > 0: valor_op = qtd * preco_unit
    if tk not in carteira: carteira[tk] = {'qtd': 0, 'custo': 0.0, 'pm': 0.0}

    # Tratamento de Proventos
    if mov in ['Rendimento', 'Juros Sobre Capital Próprio', 'Jcp', 'Dividendo']:
        proventos.append({
            'Data': row['Data'], 'AnoMes': row['Data'].strftime('%Y-%m') if pd.notna(row['Data']) else 'N/D',
            'Ticker': tk, 'Valor': valor_op, 'Tipo_Prov': mov
        })
        continue

    # Operações de Capital
    if mov == 'Compra' and es == 'Credito':
        carteira[tk]['qtd'] += qtd
        carteira[tk]['custo'] += valor_op
    elif mov in ['Venda'] and es == 'Debito':
        carteira[tk]['qtd'] -= qtd
        carteira[tk]['custo'] = carteira[tk]['qtd'] * carteira[tk]['pm'] if carteira[tk]['qtd'] > 0 else 0.0
    elif mov == 'Desdobro' and es == 'Credito':
        carteira[tk]['qtd'] += qtd
    elif mov == 'Atualização' and es == 'Credito':
        carteira[tk]['qtd'] += qtd
        if valor_op > 0:
            carteira[tk]['custo'] += valor_op
        else:
            tk_antigo = ticker_origem.get(tk)
            if tk_antigo and tk_antigo in carteira:
                carteira[tk]['custo'] += (qtd * carteira[tk_antigo]['pm'])
                carteira[tk_antigo]['qtd'] -= qtd
                carteira[tk_antigo]['custo'] = carteira[tk_antigo]['qtd'] * carteira[tk_antigo]['pm'] if carteira[tk_antigo]['qtd'] > 0 else 0

    carteira[tk]['qtd'] = max(0, carteira[tk]['qtd'])
    if carteira[tk]['qtd'] > 0:
        carteira[tk]['pm'] = carteira[tk]['custo'] / carteira[tk]['qtd']
    else:
        carteira[tk]['pm'] = 0.0

    if pd.notna(row['Data']):
        historico_pm.append({'Data': row['Data'], 'AnoMes': row['Data'].strftime('%Y-%m'), 'Ticker': tk, 'Qtd': carteira[tk]['qtd'], 'PM': carteira[tk]['pm'], 'Investido': carteira[tk]['custo']})

df_prov = pd.DataFrame(proventos)

# ==============================================================================
# 4. EXTRATOR DE DADOS DE MERCADO (YFINANCE) E INF_ATIVOS
# ==============================================================================
print("🌐 Sincronizando com o Mercado (Yahoo Finance e Google Finance)...")
ativos_ativos = [tk for tk, dados in carteira.items() if dados['qtd'] > 0 and tk not in conversao_tickers]

# Atualiza aba Inf_Ativos
try: aba_inf = planilha.worksheet("Inf_Ativos")
except: aba_inf = planilha.add_worksheet(title="Inf_Ativos", rows="1000", cols="20")
df_inf_sheet = pd.DataFrame(aba_inf.get_all_records(value_render_option='UNFORMATTED_VALUE'))
mapa_tipos = dict(zip(df_inf_sheet.iloc[:, 0].astype(str), df_inf_sheet.iloc[:, 2])) if not df_inf_sheet.empty else {}
mapa_segs = dict(zip(df_inf_sheet.iloc[:, 0].astype(str), df_inf_sheet.iloc[:, 3])) if not df_inf_sheet.empty else {}

# Consolidando Resultados Atuais
resultados = []
for tk in ativos_ativos:
    resultados.append({
        'Ticker': tk, 'Qtd': carteira[tk]['qtd'], 'PM': carteira[tk]['pm'], 'Investido': carteira[tk]['custo'],
        'Tipo': mapa_tipos.get(tk, 'Outros'), 'Seguimento': mapa_segs.get(tk, 'Outros')
    })
df_res = pd.DataFrame(resultados).sort_values(by='Investido', ascending=False)
df_res['Cotação'] = df_res['Ticker'].apply(lambda x: f'=IFERROR(GOOGLEFINANCE("BVMF:{x}"; "price"); 0)')
df_res['Patrimônio'] = [f'=B{i+2}*E{i+2}' for i in range(len(df_res))]
df_res['Variação'] = [f'=IFERROR((F{i+2}/D{i+2})-1; 0)' for i in range(len(df_res))]

try: aba_carteira = planilha.worksheet("Carteira")
except: aba_carteira = planilha.add_worksheet(title="Carteira", rows="100", cols="20")
aba_carteira.batch_clear(["A:H"])
aba_carteira.update(range_name='A1', values=[df_res.columns.tolist()] + df_res.values.tolist(), value_input_option='USER_ENTERED')

# ==============================================================================
# 5. INTELIGÊNCIA E CÁLCULO DE INSIGHTS GLOBAIS
# ==============================================================================
print("🧠 Processando Insights de Investimento...")
time.sleep(3) # Aguarda Google Finance processar
df_res_atualizado = pd.DataFrame(aba_carteira.get_all_records(value_render_option='UNFORMATTED_VALUE'))

total_investido = df_res_atualizado['Investido'].sum()
total_patrimonio = df_res_atualizado['Patrimônio'].sum()
total_proventos = df_prov['Valor'].sum() if not df_prov.empty else 0.0

data_12m = datetime.now() - relativedelta(years=1)
df_prov_12m = df_prov[df_prov['Data'] >= data_12m] if not df_prov.empty else pd.DataFrame()
prov_12m = df_prov_12m['Valor'].sum() if not df_prov_12m.empty else 0.0

lucro_capital = total_patrimonio - total_investido
lucro_total = lucro_capital + total_proventos
yoc_12m = (prov_12m / total_investido * 100) if total_investido > 0 else 0

# Insights Automáticos Gerados por IA Simbólica
insights = []
if lucro_capital > 0: insights.append(f"🟢 Seu patrimônio teve uma valorização real de mercado de **R$ {lucro_capital:,.2f}**.")
else: insights.append(f"🔴 Sua carteira está descontada em **R$ {abs(lucro_capital):,.2f}**. Boa oportunidade para aportes?")

if yoc_12m > 6.0: insights.append(f"🔥 Excelente! Seu *Yield on Cost* (12M) está em **{yoc_12m:.2f}%**, batendo a inflação com folga.")
else: insights.append(f"💡 Seu *Yield on Cost* (12M) é de **{yoc_12m:.2f}%**. Foque em ativos pagadores para gerar mais renda passiva.")

if not df_res_atualizado.empty:
    campeao = df_res_atualizado.loc[df_res_atualizado['Variação'].idxmax()]
    insights.append(f"🏆 O ativo campeão da carteira é **{campeao['Ticker']}**, com **{campeao['Variação']*100:.2f}%** de lucro acumulado.")

# ==============================================================================
# 6. GERAÇÃO DO DASHBOARD PREMIUM (HTML/JS)
# ==============================================================================
print("🎨 Gerando Dashboard Premium...")

# Preparando dados para o JS
js_carteira = df_res_atualizado.to_json(orient='records')
js_proventos = df_prov.to_json(orient='records') if not df_prov.empty else "[]"
js_insights = json.dumps(insights)

html_template = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Portfólio Intelligence</title>
    <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
        
        :root {
            --bg: #0f172a; --surface: #1e293b; --surface-hover: #334155;
            --text: #f8fafc; --text-muted: #94a3b8;
            --primary: #3b82f6; --accent: #8b5cf6;
            --success: #10b981; --danger: #ef4444; --warning: #f59e0b;
            --border: rgba(255,255,255,0.1);
        }

        * { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Inter', sans-serif; }
        body { background: var(--bg); color: var(--text); padding: 20px; min-height: 100vh; }
        
        /* Glassmorphism Header */
        header {
            background: rgba(30, 41, 59, 0.7); backdrop-filter: blur(10px);
            border-bottom: 1px solid var(--border);
            padding: 20px 30px; border-radius: 16px; margin-bottom: 24px;
            display: flex; justify-content: space-between; align-items: center;
        }
        .logo { font-size: 20px; font-weight: 700; background: linear-gradient(to right, var(--primary), var(--accent)); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
        
        .nav-pills { display: flex; gap: 10px; background: rgba(0,0,0,0.2); padding: 5px; border-radius: 12px; }
        .nav-pills button {
            background: transparent; border: none; color: var(--text-muted); padding: 8px 16px; border-radius: 8px; font-weight: 500; cursor: pointer; transition: 0.3s;
        }
        .nav-pills button.active { background: var(--primary); color: white; box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3); }

        /* Grid System */
        .dashboard-container { display: none; animation: fadeIn 0.4s ease; }
        .dashboard-container.active { display: block; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }

        .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 20px; margin-bottom: 24px; }
        
        /* Cards */
        .card {
            background: var(--surface); border: 1px solid var(--border); border-radius: 16px;
            padding: 24px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); transition: transform 0.2s;
        }
        .card:hover { transform: translateY(-2px); border-color: rgba(255,255,255,0.2); }
        
        .kpi-header { display: flex; justify-content: space-between; color: var(--text-muted); font-size: 14px; font-weight: 500; margin-bottom: 12px; }
        .kpi-value { font-size: 28px; font-weight: 700; margin-bottom: 8px; }
        .kpi-badge { font-size: 12px; padding: 4px 8px; border-radius: 6px; font-weight: 600; }
        .bg-success { background: rgba(16, 185, 129, 0.2); color: var(--success); }
        .bg-danger { background: rgba(239, 68, 68, 0.2); color: var(--danger); }

        .charts-grid { display: grid; grid-template-columns: 2fr 1fr; gap: 20px; margin-bottom: 24px; }
        .chart-container { width: 100%; height: 350px; }
        
        /* AI Insights Panel */
        .insights-panel { background: linear-gradient(145deg, rgba(59,130,246,0.1) 0%, rgba(139,92,246,0.1) 100%); border-left: 4px solid var(--accent); padding: 20px; border-radius: 0 16px 16px 0; margin-bottom: 24px; }
        .insights-panel h3 { margin-bottom: 12px; font-size: 16px; display: flex; align-items: center; gap: 8px;}
        .insights-list { list-style: none; }
        .insights-list li { margin-bottom: 8px; font-size: 14px; color: #cbd5e1; line-height: 1.5; }

        @media (max-width: 900px) { .charts-grid { grid-template-columns: 1fr; } header { flex-direction: column; gap: 15px;} }
    </style>
</head>
<body>

    <header>
        <div class="logo"><i class="fa-solid fa-chart-pie"></i> Portfólio Intelligence</div>
        <div class="nav-pills">
            <button class="active" onclick="switchTab('resumo')">Visão Geral</button>
            <button onclick="switchTab('renda')">Máquina de Renda</button>
        </div>
    </header>

    <!-- ABA 1: VISÃO GERAL -->
    <div id="resumo" class="dashboard-container active">
        
        <div class="insights-panel">
            <h3><i class="fa-solid fa-wand-magic-sparkles" style="color: var(--accent);"></i> Insights Gerados por IA</h3>
            <ul class="insights-list" id="ai-insights"></ul>
        </div>

        <div class="kpi-grid">
            <div class="card">
                <div class="kpi-header"><span>Patrimônio Atual</span><i class="fa-solid fa-wallet"></i></div>
                <div class="kpi-value">__PATRIMONIO__</div>
                <div><span class="kpi-badge __COR_VAR__">__VAR_PCT__%</span> <span style="color: var(--text-muted); font-size: 12px;">vs Investido (__INVESTIDO__)</span></div>
            </div>
            <div class="card">
                <div class="kpi-header"><span>Lucro Total (Capital + Divs)</span><i class="fa-solid fa-arrow-trend-up"></i></div>
                <div class="kpi-value __COR_TEXT_LUCRO__">__LUCRO_TOTAL__</div>
                <div style="color: var(--text-muted); font-size: 12px;">Capital: __LUCRO_CAPITAL__ | Proventos: __PROVENTOS_TOTAL__</div>
            </div>
            <div class="card">
                <div class="kpi-header"><span>Yield on Cost (12M)</span><i class="fa-solid fa-money-bill-wave"></i></div>
                <div class="kpi-value text-primary">__YOC_12M__%</div>
                <div style="color: var(--text-muted); font-size: 12px;">O quanto seu custo histórico rende hoje.</div>
            </div>
        </div>

        <div class="charts-grid">
            <div class="card">
                <h3 style="margin-bottom: 15px; font-size: 16px;">Jornada do Patrimônio (Waterfall)</h3>
                <div id="chart-waterfall" class="chart-container"></div>
            </div>
            <div class="card">
                <h3 style="margin-bottom: 15px; font-size: 16px;">Alocação por Setor</h3>
                <div id="chart-donut" class="chart-container"></div>
            </div>
        </div>
    </div>

    <!-- ABA 2: MÁQUINA DE RENDA -->
    <div id="renda" class="dashboard-container">
        <div class="card" style="margin-bottom: 20px;">
            <h3 style="margin-bottom: 15px; font-size: 16px;">Evolução de Proventos Mensais</h3>
            <div id="chart-bar-prov" class="chart-container" style="height: 450px;"></div>
        </div>
    </div>

    <script>
        const formatBRL = (val) => new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(val);
        
        function switchTab(tabId) {
            document.querySelectorAll('.dashboard-container').forEach(e => e.classList.remove('active'));
            document.querySelectorAll('.nav-pills button').forEach(e => e.classList.remove('active'));
            document.getElementById(tabId).classList.add('active');
            event.target.classList.add('active');
            window.dispatchEvent(new Event('resize'));
        }

        // Carregar Insights
        const insights = __JSON_INSIGHTS__;
        const ulInsights = document.getElementById('ai-insights');
        insights.forEach(text => {
            let li = document.createElement('li');
            li.innerHTML = text;
            ulInsights.appendChild(li);
        });

        // Configuração Padrão de Gráficos (Dark Theme)
        const darkLayout = {
            paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: 'rgba(0,0,0,0)',
            font: { color: '#94a3b8', family: 'Inter' },
            margin: { t: 20, l: 40, r: 20, b: 40 },
            xaxis: { gridcolor: 'rgba(255,255,255,0.05)', zerolinecolor: 'rgba(255,255,255,0.1)' },
            yaxis: { gridcolor: 'rgba(255,255,255,0.05)', zerolinecolor: 'rgba(255,255,255,0.1)' }
        };

        // 1. Gráfico Waterfall (Cascata)
        const valInvestido = __VAL_INVESTIDO__;
        const valGanho = __VAL_GANHO__;
        const valProv = __VAL_PROV__;
        const valFinal = valInvestido + valGanho + valProv;

        const traceWaterfall = {
            type: "waterfall", orientation: "v",
            measure: ["absolute", "relative", "relative", "total"],
            x: ["Total Investido", "Ganho de Capital", "Dividendos", "Resultado Final"],
            y: [valInvestido, valGanho, valProv, valFinal],
            textposition: "outside",
            text: [formatBRL(valInvestido), formatBRL(valGanho), formatBRL(valProv), formatBRL(valFinal)],
            decreasing: { marker: { color: "#ef4444" } },
            increasing: { marker: { color: "#10b981" } },
            totals: { marker: { color: "#3b82f6" } }
        };
        Plotly.newPlot('chart-waterfall', [traceWaterfall], {...darkLayout, hovermode: 'closest'}, {responsive: true, displayModeBar: false});

        // 2. Gráfico Donut (Alocação)
        const carteira = __JSON_CARTEIRA__;
        let setores = {};
        carteira.forEach(c => { setores[c.Seguimento] = (setores[c.Seguimento] || 0) + c['Patrimônio']; });
        
        const traceDonut = {
            values: Object.values(setores), labels: Object.keys(setores),
            type: 'pie', hole: .6,
            marker: { colors: ['#3b82f6', '#8b5cf6', '#10b981', '#f59e0b', '#ec4899'] },
            textinfo: 'label+percent', hoverinfo: 'label+value',
            textfont: { color: '#fff' }
        };
        Plotly.newPlot('chart-donut', [traceDonut], {...darkLayout, showlegend: false}, {responsive: true, displayModeBar: false});

        // 3. Gráfico de Barras Proventos
        const proventos = __JSON_PROVENTOS__;
        let provMensal = {};
        proventos.forEach(p => { provMensal[p.AnoMes] = (provMensal[p.AnoMes] || 0) + p.Valor; });
        
        const mesesSort = Object.keys(provMensal).sort();
        const valsProv = mesesSort.map(m => provMensal[m]);

        const traceBarProv = {
            x: mesesSort, y: valsProv, type: 'bar',
            marker: { color: '#10b981', borderRadius: 4 },
            text: valsProv.map(v => formatBRL(v)), textposition: 'outside'
        };
        Plotly.newPlot('chart-bar-prov', [traceBarProv], darkLayout, {responsive: true, displayModeBar: false});

    </script>
</body>
</html>"""

# Substituições HTML
def fmoeda(v): return f"R$ {v:,.2f}".replace(',','X').replace('.',',').replace('X','.')

var_pct = (lucro_capital / total_investido * 100) if total_investido > 0 else 0
cor_var = "bg-success" if var_pct >= 0 else "bg-danger"
cor_txt_lucro = "color: var(--success);" if lucro_total >= 0 else "color: var(--danger);"

html = html_template.replace('__PATRIMONIO__', fmoeda(total_patrimonio))
html = html.replace('__INVESTIDO__', fmoeda(total_investido))
html = html.replace('__LUCRO_TOTAL__', fmoeda(lucro_total))
html = html.replace('__LUCRO_CAPITAL__', fmoeda(lucro_capital))
html = html.replace('__PROVENTOS_TOTAL__', fmoeda(total_proventos))
html = html.replace('__VAR_PCT__', f"{var_pct:.2f}".replace('.',','))
html = html.replace('__COR_VAR__', cor_var)
html = html.replace('__COR_TEXT_LUCRO__', cor_txt_lucro)
html = html.replace('__YOC_12M__', f"{yoc_12m:.2f}".replace('.',','))

# Variaveis para os Gráficos
html = html.replace('__VAL_INVESTIDO__', str(total_investido))
html = html.replace('__VAL_GANHO__', str(lucro_capital))
html = html.replace('__VAL_PROV__', str(total_proventos))
html = html.replace('__JSON_CARTEIRA__', js_carteira)
html = html.replace('__JSON_PROVENTOS__', js_proventos)
html = html.replace('__JSON_INSIGHTS__', js_insights)

with open("dashboard_premium.html", "w", encoding="utf-8") as f: f.write(html)
print("✅ Sucesso! Dashboard 'dashboard_premium.html' gerado.")
