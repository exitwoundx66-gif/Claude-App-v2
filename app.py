"""
Allocation Station — personal portfolio + debt + margin command center.
Streamlit Community Cloud app. Data persists in a private Google Sheet.
"""
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, date
import json

st.set_page_config(page_title="Allocation Station", page_icon="●",
                   layout="wide", initial_sidebar_state="collapsed")

# ============================================================ CONFIG
TICKERS = [
    ("IBIT", "Core",   "#f5a524"), ("VGT",  "Core",   "#58a6ff"),
    ("SPMO", "Core",   "#4ac2a8"), ("TSLA", "Core",   "#e85d75"),
    ("SPCX", "Core",   "#a371f7"), ("BTCI", "Income", "#d4a017"),
    ("SCHD", "Income", "#6ea8fe"), ("QQQI", "Income", "#57caab"),
]
TK = [t[0] for t in TICKERS]
SLEEVE = {t[0]: t[1] for t in TICKERS}
COLOR = {t[0]: t[2] for t in TICKERS}
DEFAULT_PCT = {"IBIT":15.0,"VGT":12.5,"SPMO":12.5,"TSLA":10.0,"SPCX":10.0,
               "BTCI":13.33,"SCHD":13.33,"QQQI":13.34}
FALLBACK_PX = {"IBIT":33.04,"VGT":118.81,"SPMO":161.03,"TSLA":440.0,"SPCX":0.0,
               "BTCI":26.61,"SCHD":32.50,"QQQI":56.74}
# rough annual distribution yields for income forecast
YIELD = {"BTCI":0.27,"SCHD":0.032,"QQQI":0.135,"VGT":0.006,"SPMO":0.006,
         "IBIT":0.0,"TSLA":0.0,"SPCX":0.0}
DRIP_TARGET_DEFAULT = 4000.0

# ============================================================ STYLE
st.markdown("""<style>
:root{--amber:#f5a524;}
.stApp{background:#0d1117;}
section.main>div{padding-top:1rem;}
h1,h2,h3,h4{color:#e8edf5;font-family:'Inter',sans-serif;}
[data-testid="stMetricValue"]{font-family:'SF Mono',monospace;font-size:1.6rem;}
[data-testid="stMetricLabel"]{color:#8b98ad;text-transform:uppercase;font-size:.7rem;letter-spacing:.08em;}
.stTabs [data-baseweb="tab-list"]{gap:2px;background:#0d1117;}
.stTabs [data-baseweb="tab"]{color:#8b98ad;font-size:.85rem;padding:8px 14px;}
.stTabs [aria-selected="true"]{color:#f5a524;border-bottom-color:#f5a524;}
.block-container{max-width:1100px;padding-top:1.5rem;}
div[data-testid="stDataFrame"]{border:1px solid #2a3346;border-radius:8px;}
.badge{display:inline-block;padding:3px 10px;border-radius:5px;font-size:.75rem;
  font-family:monospace;font-weight:600;}
.b-green{background:rgba(63,185,80,.15);color:#3fb950;}
.b-amber{background:rgba(245,165,36,.15);color:#f5a524;}
.b-red{background:rgba(248,81,73,.16);color:#f85149;}
.hint{color:#5a6779;font-size:.8rem;font-style:italic;}
.brandbar{display:flex;align-items:center;gap:10px;margin-bottom:.5rem;}
.brandbar .dot{width:9px;height:9px;border-radius:50%;background:#f5a524;box-shadow:0 0 10px #f5a524;}
.brandbar h1{font-size:1.3rem;margin:0;}
.brandbar .sub{margin-left:auto;color:#5a6779;font-family:monospace;font-size:.75rem;}
</style>""", unsafe_allow_html=True)

# ============================================================ AUTH
def check_password():
    if "app_password" not in st.secrets:
        return True  # no password set -> open (local dev)
    def entered():
        st.session_state.auth = (st.session_state.pw == st.secrets["app_password"])
    if st.session_state.get("auth"):
        return True
    st.markdown("### ● Allocation Station")
    st.text_input("Password", type="password", key="pw", on_change=entered)
    if st.session_state.get("auth") is False:
        st.error("Incorrect password.")
    return False

if not check_password():
    st.stop()

# ============================================================ DATA LAYER (Google Sheets)
import gspread
from google.oauth2.service_account import Credentials

SHEETS = {
    "buys":     ["date","ticker","amount","price"],
    "divs":     ["date","ticker","amount","dest"],
    "snaps":    ["month","value","contrib","loan"],
    "debts":    ["name","balance","apr","min_pmt"],
    "config":   ["key","value"],
}

@st.cache_resource
def connect():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open_by_key(st.secrets["sheet_key"])

def ws(name):
    sh = connect()
    try:
        w = sh.worksheet(name)
    except gspread.WorksheetNotFound:
        w = sh.add_worksheet(name, rows=200, cols=max(4,len(SHEETS[name])))
        w.append_row(SHEETS[name])
    return w

@st.cache_data(ttl=30)
def load(name):
    try:
        recs = ws(name).get_all_records()
        df = pd.DataFrame(recs)
        if df.empty:
            df = pd.DataFrame(columns=SHEETS[name])
        return df
    except Exception as e:
        st.session_state["_err"] = str(e)
        return pd.DataFrame(columns=SHEETS[name])

def append(name, row):
    ws(name).append_row(row)
    load.clear()

def rewrite(name, df):
    w = ws(name)
    w.clear()
    w.append_row(SHEETS[name])
    if not df.empty:
        w.append_rows(df.astype(object).values.tolist())
    load.clear()

def cfg_get(key, default):
    df = load("config")
    if not df.empty and key in df["key"].values:
        v = df.loc[df["key"]==key,"value"].iloc[0]
        try: return float(v)
        except: return v
    return default

def cfg_set(key, value):
    df = load("config")
    if not df.empty and key in df["key"].values:
        df.loc[df["key"]==key,"value"] = value
    else:
        df = pd.concat([df, pd.DataFrame([{"key":key,"value":value}])], ignore_index=True)
    rewrite("config", df)

# ============================================================ LIVE PRICES
@st.cache_data(ttl=900)
def fetch_prices(tickers):
    import yfinance as yf
    out = {}
    try:
        data = yf.download(tickers, period="1d", progress=False, threads=True)
        closes = data["Close"] if "Close" in data else data
        for t in tickers:
            try:
                v = float(closes[t].dropna().iloc[-1])
                out[t] = v if v>0 else None
            except Exception:
                out[t] = None
    except Exception:
        for t in tickers: out[t] = None
    return out

def get_prices():
    live = fetch_prices(TK)
    prices, sources = {}, {}
    saved = json.loads(cfg_get("saved_prices", json.dumps(FALLBACK_PX)))
    for t in TK:
        if live.get(t):
            prices[t] = live[t]; sources[t] = "live"
        elif saved.get(t):
            prices[t] = saved[t]; sources[t] = "saved"
        else:
            prices[t] = FALLBACK_PX.get(t,0); sources[t] = "default"
    # persist any live prices as the new fallback
    if any(v=="live" for v in sources.values()):
        merged = {**saved, **{t:prices[t] for t in TK if sources[t]=="live"}}
        try: cfg_set("saved_prices", json.dumps(merged))
        except Exception: pass
    return prices, sources

# ============================================================ COMPUTE
def get_pct():
    raw = cfg_get("pct", json.dumps(DEFAULT_PCT))
    try: return {**DEFAULT_PCT, **json.loads(raw)}
    except: return dict(DEFAULT_PCT)

def holdings_df(prices):
    buys = load("buys")
    rows = []
    for t in TK:
        b = buys[buys["ticker"]==t] if not buys.empty else pd.DataFrame()
        inv = float(b["amount"].astype(float).sum()) if not b.empty else 0.0
        sh = float((b["amount"].astype(float)/b["price"].astype(float).replace(0,pd.NA)).sum()) if not b.empty else 0.0
        sh = 0.0 if pd.isna(sh) else sh
        px = prices.get(t,0)
        val = sh*px
        rows.append(dict(Ticker=t, Sleeve=SLEEVE[t], Invested=inv, Shares=sh,
                         Avg=(inv/sh if sh>0 else 0), Price=px, Value=val,
                         PL=val-inv, PLpct=((val-inv)/inv if inv>0 else 0)))
    return pd.DataFrame(rows)

def totals(h):
    return dict(inv=h["Invested"].sum(), val=h["Value"].sum(),
                pl=h["Value"].sum()-h["Invested"].sum(),
                core=h[h["Sleeve"]=="Core"]["Value"].sum(),
                income=h[h["Sleeve"]=="Income"]["Value"].sum())

def money(n): 
    try: return f"${n:,.2f}"
    except: return "—"
def money0(n):
    try: return f"${n:,.0f}"
    except: return "—"

# ============================================================ HEADER
prices, sources = get_prices()
h = holdings_df(prices)
T = totals(h)
nlive = sum(1 for s in sources.values() if s=="live")
st.markdown(f"""<div class='brandbar'><span class='dot'></span>
<h1>Allocation Station</h1>
<span class='sub'>{nlive}/{len(TK)} live prices · {datetime.now():%b %d %H:%M}</span></div>""",
unsafe_allow_html=True)

tabs = st.tabs(["Dashboard","Invest","Holdings","Income","Growth","Margin","Debt","Goal"])

# ============================================================ TAB: DASHBOARD
with tabs[0]:
    c = st.columns(4)
    c[0].metric("Invested", money0(T["inv"]))
    c[1].metric("Value", money0(T["val"]))
    c[2].metric("Gain / Loss", money0(T["pl"]), f"{(T['pl']/T['inv']*100 if T['inv'] else 0):+.1f}%")
    drip_target = cfg_get("drip_target", DRIP_TARGET_DEFAULT)
    prog = min(1, T["val"]/drip_target) if drip_target else 0
    c[3].metric("DRIP → $%s" % f"{drip_target:,.0f}", f"{prog*100:.0f}%",
                "ROUTE TO MARGIN" if T["val"]>=drip_target else "keep DRIP on")

    left, right = st.columns([1,1])
    with left:
        st.markdown("**Allocation**")
        vals = [max(0,r.Value) for r in h.itertuples()]
        if sum(vals)>0:
            fig = go.Figure(go.Pie(labels=TK, values=vals, hole=.62,
                marker=dict(colors=[COLOR[t] for t in TK]),
                textinfo="none", hoverinfo="label+percent"))
            fig.update_layout(showlegend=True, height=280, margin=dict(t=10,b=10,l=0,r=0),
                paper_bgcolor="rgba(0,0,0,0)", font=dict(color="#e8edf5",size=11),
                legend=dict(font=dict(size=10)),
                annotations=[dict(text=money0(T["val"]),x=.5,y=.5,font_size=15,showarrow=False,
                    font_color="#e8edf5")])
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})
        else:
            st.info("Log your first buy on the Holdings tab to populate this.")
    with right:
        st.markdown("**Drift vs target**")
        pct = get_pct()
        drows = []
        for r in h.itertuples():
            actual = (r.Value/T["val"]) if T["val"]>0 else 0
            target = pct.get(r.Ticker,0)/100
            drows.append(dict(Ticker=r.Ticker, Target=f"{target*100:.1f}%",
                              Actual=f"{actual*100:.1f}%", Drift=f"{(actual-target)*100:+.1f}%"))
        st.dataframe(pd.DataFrame(drows), hide_index=True, use_container_width=True, height=320)
    st.markdown("<span class='hint'>DRIP distributions until the portfolio clears the target, "
                "then route dividends to margin paydown.</span>", unsafe_allow_html=True)

# ============================================================ TAB: INVEST (calculator + rebalance)
with tabs[1]:
    st.markdown("#### Invest calculator")
    amt = st.number_input("Amount to invest ($)", min_value=0.0, value=1000.0, step=50.0)
    mode = st.radio("Split method", ["By target %","Rebalance-aware (fill underweights first)"],
                    horizontal=True)
    pct = get_pct()
    if mode.startswith("By target"):
        split = {t: amt*pct[t]/100 for t in TK}
    else:
        # rebalance-aware: bring the most-underweight holdings up toward target first
        future_total = T["val"] + amt
        need = {t: max(0, future_total*pct[t]/100 - h.loc[h.Ticker==t,"Value"].iloc[0]) for t in TK}
        tot_need = sum(need.values())
        if tot_need <= amt and tot_need > 0:
            # fill all gaps, distribute remainder by target %
            rem = amt - tot_need
            split = {t: need[t] + rem*pct[t]/100 for t in TK}
        elif tot_need > 0:
            split = {t: amt*need[t]/tot_need for t in TK}  # proportional to gap
        else:
            split = {t: amt*pct[t]/100 for t in TK}
    dfc = pd.DataFrame([dict(Ticker=t, Sleeve=SLEEVE[t], Target=f"{pct[t]:.2f}%",
                             Buy=round(split[t],2)) for t in TK])
    st.dataframe(dfc, hide_index=True, use_container_width=True,
                 column_config={"Buy":st.column_config.NumberColumn("Buy $", format="$%.2f")})
    cc = st.columns(3)
    cc[0].metric("Core", money(sum(split[t] for t in TK if SLEEVE[t]=="Core")))
    cc[1].metric("Income", money(sum(split[t] for t in TK if SLEEVE[t]=="Income")))
    cc[2].metric("Total", money(sum(split.values())))
    st.markdown("<span class='hint'>Rebalance-aware mode steers new money toward whatever has "
                "fallen furthest below target — no selling, just smarter buys.</span>", unsafe_allow_html=True)

    with st.expander("Edit target allocation %"):
        newp = {}
        cols = st.columns(4)
        for i,t in enumerate(TK):
            newp[t] = cols[i%4].number_input(t, value=float(pct[t]), step=0.5, key=f"p_{t}")
        s = sum(newp.values())
        st.write(f"Sum: **{s:.2f}%** " + ("✓" if abs(s-100)<0.01 else "⚠ must equal 100%"))
        if st.button("Save allocation"):
            cfg_set("pct", json.dumps(newp)); st.success("Saved."); st.rerun()

# ============================================================ TAB: HOLDINGS (+ buy log + prices)
with tabs[2]:
    st.markdown("#### Log a buy")
    bc = st.columns([1.2,1,1,1])
    bt = bc[0].selectbox("Ticker", TK, key="bt")
    ba = bc[1].number_input("Amount $", min_value=0.0, step=10.0, key="ba")
    bp = bc[2].number_input("Price paid", min_value=0.0, value=float(prices.get(bt,0)), step=0.5, key="bp")
    bc[3].write(""); bc[3].write("")
    if bc[3].button("Add buy", use_container_width=True):
        if ba>0 and bp>0:
            append("buys",[date.today().isoformat(),bt,ba,bp]); st.success(f"Added {money(ba)} {bt}"); st.rerun()
        else: st.warning("Enter amount and price.")
    if ba>0 and bp>0:
        st.caption(f"= {ba/bp:.4f} shares")

    st.markdown("#### Holdings")
    disp = h.copy()
    disp = disp[disp["Invested"]>0] if not h.empty else h
    show = h[["Ticker","Invested","Shares","Avg","Price","Value","PL","PLpct"]].copy()
    st.dataframe(show, hide_index=True, use_container_width=True, column_config={
        "Invested":st.column_config.NumberColumn(format="$%.2f"),
        "Shares":st.column_config.NumberColumn(format="%.4f"),
        "Avg":st.column_config.NumberColumn(format="$%.2f"),
        "Price":st.column_config.NumberColumn(format="$%.2f"),
        "Value":st.column_config.NumberColumn(format="$%.2f"),
        "PL":st.column_config.NumberColumn("P/L", format="$%.2f"),
        "PLpct":st.column_config.NumberColumn("P/L %", format="%.1f%%")})
    m = st.columns(3)
    m[0].metric("Total invested", money0(T["inv"]))
    m[1].metric("Total value", money0(T["val"]))
    m[2].metric("Total P/L", money0(T["pl"]))

    with st.expander("Manual price override (used when live fetch is unavailable)"):
        saved = json.loads(cfg_get("saved_prices", json.dumps(FALLBACK_PX)))
        cols = st.columns(4); newpx={}
        for i,t in enumerate(TK):
            newpx[t] = cols[i%4].number_input(f"{t} ({sources[t]})", value=float(prices[t]), step=0.5, key=f"px_{t}")
        if st.button("Save manual prices"):
            cfg_set("saved_prices", json.dumps({**saved,**newpx})); fetch_prices.clear()
            st.success("Saved. Live prices still take priority when available."); st.rerun()

    buys = load("buys")
    if not buys.empty:
        with st.expander(f"Buy history ({len(buys)} rows) — delete entries"):
            b2 = buys.copy().reset_index(drop=True)
            b2["shares"] = (b2["amount"].astype(float)/b2["price"].astype(float)).round(4)
            st.dataframe(b2, hide_index=True, use_container_width=True)
            idx = st.number_input("Row # to delete (0-based)", min_value=0,
                                  max_value=max(0,len(b2)-1), step=1)
            if st.button("Delete row"):
                b3 = buys.drop(buys.index[int(idx)]).reset_index(drop=True)
                rewrite("buys", b3); st.success("Deleted."); st.rerun()

# ============================================================ TAB: INCOME
with tabs[3]:
    st.markdown("#### Log distribution")
    dc = st.columns([1.2,1,1,1])
    dt = dc[0].selectbox("Ticker", TK, key="dt")
    da = dc[1].number_input("Amount $", min_value=0.0, step=1.0, key="da")
    dd = dc[2].selectbox("Destination", ["DRIP","MARGIN"], key="dd")
    dc[3].write(""); dc[3].write("")
    if dc[3].button("Log distribution", use_container_width=True):
        if da>0:
            append("divs",[date.today().isoformat(),dt,da,dd]); st.success(f"Logged {money(da)} {dt}"); st.rerun()
        else: st.warning("Enter amount.")

    divs = load("divs")
    dripd = float(divs[divs["dest"]=="DRIP"]["amount"].astype(float).sum()) if not divs.empty else 0
    margd = float(divs[divs["dest"]=="MARGIN"]["amount"].astype(float).sum()) if not divs.empty else 0
    mc = st.columns(3)
    mc[0].metric("Reinvested (DRIP)", money0(dripd))
    mc[1].metric("To margin", money0(margd))
    mc[2].metric("Total collected", money0(dripd+margd))

    st.markdown("#### Forward income forecast")
    fc_rows=[]; annual=0
    for r in h.itertuples():
        inc = r.Value*YIELD.get(r.Ticker,0)
        annual += inc
        if r.Value>0 and YIELD.get(r.Ticker,0)>0:
            fc_rows.append(dict(Ticker=r.Ticker, Value=round(r.Value,2),
                                Yield=f"{YIELD[r.Ticker]*100:.1f}%",
                                Annual=round(inc,2), Monthly=round(inc/12,2)))
    if fc_rows:
        st.dataframe(pd.DataFrame(fc_rows), hide_index=True, use_container_width=True, column_config={
            "Value":st.column_config.NumberColumn(format="$%.0f"),
            "Annual":st.column_config.NumberColumn("Annual $", format="$%.0f"),
            "Monthly":st.column_config.NumberColumn("Monthly $", format="$%.0f")})
    fm = st.columns(3)
    fm[0].metric("Projected annual income", money0(annual))
    fm[1].metric("Projected monthly", money0(annual/12))
    fm[2].metric("Blended yield", f"{(annual/T['val']*100 if T['val'] else 0):.2f}%")
    st.markdown("<span class='hint'>Forecast uses rough current distribution rates "
                "(BTCI 27%, QQQI 13.5%, SCHD 3.2%). Actual payouts float — verify quarterly.</span>",
                unsafe_allow_html=True)

    if not divs.empty:
        with st.expander(f"Distribution history ({len(divs)} rows)"):
            st.dataframe(divs, hide_index=True, use_container_width=True)
            idx = st.number_input("Row # to delete", min_value=0, max_value=max(0,len(divs)-1),
                                  step=1, key="deldiv")
            if st.button("Delete distribution"):
                d3 = divs.drop(divs.index[int(idx)]).reset_index(drop=True)
                rewrite("divs", d3); st.success("Deleted."); st.rerun()

# ============================================================ TAB: GROWTH
with tabs[4]:
    st.markdown("#### Monthly snapshots")
    if st.button("Take snapshot now"):
        m = date.today().strftime("%Y-%m")
        moneyin = T["inv"]+dripd
        loan = cfg_get("margin_loan", 0)
        snaps = load("snaps")
        snaps = snaps[snaps["month"]!=m] if not snaps.empty else snaps
        snaps = pd.concat([snaps, pd.DataFrame([dict(month=m,value=round(T["val"],2),
                    contrib=round(moneyin,2),loan=loan)])], ignore_index=True)
        snaps = snaps.sort_values("month")
        rewrite("snaps", snaps); st.success(f"Snapshot saved for {m}"); st.rerun()

    snaps = load("snaps")
    if not snaps.empty:
        snaps = snaps.sort_values("month")
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=snaps["month"], y=snaps["contrib"].astype(float),
            fill="tozeroy", name="Contributed", line=dict(color="#58a6ff",width=1.5),
            fillcolor="rgba(88,166,255,.12)"))
        fig.add_trace(go.Scatter(x=snaps["month"], y=snaps["value"].astype(float),
            name="Value", line=dict(color="#f5a524",width=2.5)))
        fig.update_layout(height=280, margin=dict(t=10,b=10,l=0,r=0),
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#8b98ad",size=11), legend=dict(orientation="h",y=1.1),
            xaxis=dict(gridcolor="#1e2530"), yaxis=dict(gridcolor="#1e2530",tickformat="$,.0f"))
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})
        sd = snaps.copy()
        sd["mkt_gain"] = sd["value"].astype(float)-sd["contrib"].astype(float)
        st.dataframe(sd[["month","value","contrib","mkt_gain"]], hide_index=True,
            use_container_width=True, column_config={
            "value":st.column_config.NumberColumn("Value", format="$%.0f"),
            "contrib":st.column_config.NumberColumn("Contributed", format="$%.0f"),
            "mkt_gain":st.column_config.NumberColumn("Market gain", format="$%.0f")})
    else:
        st.info("Take your first snapshot to start charting growth over time.")

    st.markdown("#### Contributions vs market")
    moneyin = T["inv"]+dripd; mkt = T["val"]-moneyin
    g = st.columns(3)
    g[0].metric("Your money in", money0(moneyin))
    g[1].metric("Market gain/loss", money0(mkt))
    g[2].metric("% from market", f"{(mkt/T['val']*100 if T['val'] else 0):.0f}%")

# ============================================================ TAB: MARGIN
with tabs[5]:
    st.markdown("#### Margin state")
    loan = st.number_input("Loan balance $", min_value=0.0,
                           value=float(cfg_get("margin_loan",0)), step=100.0)
    ceil = st.number_input("Safe LTV ceiling %", min_value=1.0,
                           value=float(cfg_get("margin_ceil",35)), step=1.0)
    if loan!=cfg_get("margin_loan",0): cfg_set("margin_loan",loan)
    if ceil!=cfg_get("margin_ceil",35): cfg_set("margin_ceil",ceil)
    ltv = loan/T["val"] if T["val"]>0 else 0
    mm = st.columns(4)
    mm[0].metric("Portfolio", money0(T["val"]))
    mm[1].metric("Loan", money0(loan))
    mm[2].metric("LTV", f"{ltv*100:.1f}%")
    mm[3].metric("Equity", f"{(1-ltv)*100:.1f}%")
    if ltv>0.75: st.markdown("<span class='badge b-red'>MARGIN CALL RISK</span>", unsafe_allow_html=True)
    elif ltv*100>ceil: st.markdown("<span class='badge b-amber'>ABOVE TARGET — PAY DOWN</span>", unsafe_allow_html=True)
    else: st.markdown("<span class='badge b-green'>WITHIN SAFE ZONE</span>", unsafe_allow_html=True)

    st.markdown("#### Monthly flow")
    f = st.columns(3)
    dep = f[0].number_input("Deposit /mo", value=float(cfg_get("m_dep",600)), step=50.0)
    draw = f[1].number_input("Withdrawal /mo", value=float(cfg_get("m_draw",408)), step=50.0)
    apr = f[2].number_input("Gold APR %", value=float(cfg_get("m_apr",6.75)), step=0.25)
    for k,v in [("m_dep",dep),("m_draw",draw),("m_apr",apr)]:
        if v!=cfg_get(k,{"m_dep":600,"m_draw":408,"m_apr":6.75}[k]): cfg_set(k,v)
    ann_inc = sum(h.loc[h.Ticker==t,"Value"].iloc[0]*YIELD[t] for t in TK)
    mdiv = ann_inc/12
    interest = max(0,loan-1000)*(apr/100)/12
    counts = T["val"]>=cfg_get("drip_target",DRIP_TARGET_DEFAULT)
    net = draw+interest-(mdiv if counts else 0)
    ff = st.columns(3)
    ff[0].metric("Est. monthly dividends", money(mdiv), "counts" if counts else "DRIP phase")
    ff[1].metric("Est. monthly interest", money(interest))
    ff[2].metric("Net loan change /mo", money(net), "growing" if net>0 else "shrinking",
                 delta_color="inverse")
    st.markdown("<span class='hint'>Milestones from the modeling: loan stops growing near "
                "$56k portfolio, dividends cover a $410 draw near $117k.</span>", unsafe_allow_html=True)

# ============================================================ TAB: DEBT
with tabs[6]:
    st.markdown("#### Debt payoff tracker")
    debts = load("debts")
    with st.expander("Add / update a debt", expanded=debts.empty):
        d = st.columns([1.4,1,1,1])
        dn = d[0].text_input("Name", key="dn")
        db = d[1].number_input("Balance $", min_value=0.0, step=100.0, key="db")
        dapr = d[2].number_input("APR %", min_value=0.0, step=0.25, key="dapr")
        dmin = d[3].number_input("Min pmt $", min_value=0.0, step=10.0, key="dmin")
        if st.button("Save debt"):
            if dn and db>0:
                debts2 = debts[debts["name"]!=dn] if not debts.empty else debts
                debts2 = pd.concat([debts2, pd.DataFrame([dict(name=dn,balance=db,apr=dapr,min_pmt=dmin)])],
                                   ignore_index=True)
                rewrite("debts", debts2); st.success("Saved."); st.rerun()
            else: st.warning("Enter name and balance.")
    if not debts.empty:
        dd = debts.copy()
        dd["balance"]=dd["balance"].astype(float); dd["apr"]=dd["apr"].astype(float)
        dd["min_pmt"]=dd["min_pmt"].astype(float)
        dd = dd.sort_values("apr", ascending=False)
        dd["monthly_int"] = dd["balance"]*dd["apr"]/100/12
        st.dataframe(dd, hide_index=True, use_container_width=True, column_config={
            "balance":st.column_config.NumberColumn("Balance", format="$%.0f"),
            "apr":st.column_config.NumberColumn("APR", format="%.2f%%"),
            "min_pmt":st.column_config.NumberColumn("Min pmt", format="$%.0f"),
            "monthly_int":st.column_config.NumberColumn("Interest/mo", format="$%.2f")})
        tot_debt = dd["balance"].sum(); tot_int = dd["monthly_int"].sum()
        dm = st.columns(3)
        dm[0].metric("Total debt", money0(tot_debt))
        dm[1].metric("Interest bleeding /mo", money0(tot_int))
        dm[2].metric("Highest rate", f"{dd['apr'].max():.2f}%")
        st.markdown(f"<span class='hint'>Avalanche order (highest APR first): "
                    f"{' → '.join(dd['name'].tolist())}. Every extra dollar to the top of that list "
                    f"earns its APR, guaranteed.</span>", unsafe_allow_html=True)
        # extra payment simulator
        st.markdown("#### Payoff simulator")
        target = st.selectbox("Attack which debt", dd["name"].tolist())
        extra = st.number_input("Extra $/mo toward it", min_value=0.0, value=600.0, step=50.0)
        row = dd[dd["name"]==target].iloc[0]
        bal, apr_, pmt = row["balance"], row["apr"]/100/12, row["min_pmt"]+extra
        months=0; ti=0; b=bal
        while b>0.01 and months<600:
            i=b*apr_; ti+=i; b=b+i-pmt; months+=1
            if b<0: b=0
        base_m=0; bb=bal; base_pmt=row["min_pmt"] if row["min_pmt"]>0 else bal*apr_+50
        while bb>0.01 and base_m<600:
            bb=bb+bb*apr_-base_pmt; base_m+=1
            if bb<0: bb=0
        sc=st.columns(3)
        sc[0].metric("Payoff time", f"{months} mo" if months<600 else "—")
        sc[1].metric("Interest paid", money0(ti))
        sc[2].metric("Months saved vs min", f"{max(0,base_m-months)}" if base_m<600 else "—")
        if st.button("Delete this debt"):
            d3 = debts[debts["name"]!=target].reset_index(drop=True)
            rewrite("debts", d3); st.success("Deleted."); st.rerun()
    else:
        st.info("Add your debts to track payoff and run the avalanche simulator.")

# ============================================================ TAB: GOAL
with tabs[7]:
    st.markdown("#### Projection to goal")
    gc = st.columns(3)
    monthly = gc[0].number_input("Monthly contribution $", value=float(cfg_get("g_monthly",1008)), step=50.0)
    ret = gc[1].number_input("Assumed annual return %", value=float(cfg_get("g_ret",10.0)), step=1.0)
    goal = gc[2].number_input("Target portfolio $", value=float(cfg_get("g_goal",150000)), step=5000.0)
    for k,v in [("g_monthly",monthly),("g_ret",ret),("g_goal",goal)]:
        if v!=cfg_get(k,{"g_monthly":1008,"g_ret":10.0,"g_goal":150000}[k]): cfg_set(k,v)
    # project forward monthly
    r_m=(1+ret/100)**(1/12)-1
    val=T["val"]; months=0; series=[]; hit=None
    milestones={56000:"loan stops growing",117000:"dividends cover $410 draw",goal:"GOAL"}
    ms_hits={}
    while val<goal and months<600:
        val=val*(1+r_m)+monthly; months+=1
        series.append(val)
        for lvl,lbl in milestones.items():
            if lvl not in ms_hits and val>=lvl:
                ms_hits[lvl]=months
    if val>=goal and goal not in ms_hits: ms_hits[goal]=months
    yrs = months/12
    gm=st.columns(3)
    gm[0].metric("Current value", money0(T["val"]))
    gm[1].metric("Time to goal", f"{yrs:.1f} yrs" if months<600 else "60+ yrs")
    gm[2].metric("Reach date", (date.today().replace(day=1)+pd.DateOffset(months=months)).strftime("%b %Y") if months<600 else "—")
    if series:
        fig=go.Figure()
        xs=[(date.today()+pd.DateOffset(months=i)).strftime("%Y-%m") for i in range(len(series))]
        fig.add_trace(go.Scatter(x=xs,y=series,line=dict(color="#f5a524",width=2.5),name="Projected"))
        fig.add_hline(y=goal,line=dict(color="#3fb950",dash="dash"),annotation_text="goal")
        for lvl,lbl in [(56000,"loan stable"),(117000,"div-covered")]:
            if lvl<goal:
                fig.add_hline(y=lvl,line=dict(color="#58a6ff",dash="dot",width=1),
                              annotation_text=lbl,annotation_font_size=9)
        fig.update_layout(height=300,margin=dict(t=10,b=10,l=0,r=0),
            paper_bgcolor="rgba(0,0,0,0)",plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#8b98ad",size=11),showlegend=False,
            xaxis=dict(gridcolor="#1e2530",nticks=8),yaxis=dict(gridcolor="#1e2530",tickformat="$,.0f"))
        st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False})
    st.markdown("**Milestone timeline**")
    for lvl,lbl in sorted(milestones.items()):
        if lvl in ms_hits:
            when=(date.today()+pd.DateOffset(months=ms_hits[lvl])).strftime("%b %Y")
            st.markdown(f"- **{lbl}** (${lvl:,.0f}) → ~{when} ({ms_hits[lvl]/12:.1f} yrs)")
    st.markdown("<span class='hint'>Ties to the plan: the margin-income structure becomes "
                "sustainable around $125–150k, which this contribution path reaches on your 2032–2035 horizon.</span>",
                unsafe_allow_html=True)

# ============================================================ FOOTER
st.divider()
fc=st.columns([3,1])
fc[0].markdown("<span class='hint'>Data lives in your private Google Sheet. "
               "Prices refresh every 15 min from Yahoo Finance with manual fallback.</span>",
               unsafe_allow_html=True)
if fc[1].button("Refresh prices"):
    fetch_prices.clear(); load.clear(); st.rerun()
if "_err" in st.session_state:
    with st.expander("⚠ connection note"):
        st.code(st.session_state["_err"])
