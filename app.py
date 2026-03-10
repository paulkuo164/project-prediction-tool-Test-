import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objs as go
from datetime import timedelta
import calendar
import warnings
import io
import os

# ===== 🤐 系統設定 =====
warnings.filterwarnings('ignore')
np.seterr(divide='ignore', invalid='ignore')

st.set_page_config(page_title="工程金流預測儀表板", layout="wide")
st.title("🏗️ 工程進度預測與金流預估 (精確修正版)")

# ===== 🛠️ 1. 核心邏輯函式 =====
def get_month_end(dt):
    if pd.isna(dt) or dt is None: return None
    if isinstance(dt, str): dt = pd.to_datetime(dt)
    last_day = calendar.monthrange(dt.year, dt.month)[1]
    return dt.replace(day=last_day)

def get_payment_date(dt):
    if pd.isna(dt) or dt is None: return None
    target_date = (pd.to_datetime(dt).replace(day=1) + pd.DateOffset(months=2))
    return target_date.replace(day=5)

def clean_and_process(df, base_start_date=None):
    df.columns = [str(c).strip() for c in df.columns]
    date_candidates = [c for c in df.columns if any(k in c for k in ["日", "期"])]
    if not date_candidates: return None, None, None, None
    date_col = date_candidates[0]
    df[date_col] = pd.to_datetime(df[date_col].astype(str).str.replace(r"\s*\(.*?\)", "", regex=True), errors="coerce")
    df.dropna(subset=[date_col], inplace=True)
    actual_start = pd.to_datetime(base_start_date) if base_start_date is not None else df[date_col].min()
    df["天數"] = (df[date_col] - actual_start).dt.days
    
    cum_col = [c for c in df.columns if "累計" in c][0]
    df[cum_col] = pd.to_numeric(df[cum_col], errors="coerce")
    max_p = df[cum_col].max()
    df["累計_norm"] = (df[cum_col] / (max_p if max_p > 0 else 1)) * 100
    df["天數_norm"] = (df["天數"] / (df["天數"].max() if df["天數"].max() > 0 else 1)) * 100
    return df, date_col, cum_col, actual_start

# ===== 📂 2. 資料處理流程 =====
st.sidebar.header("📂 資料來源")
active_file = st.sidebar.file_uploader("上傳歷史案件 Excel", type=["xlsx"])

if active_file:
    xls = pd.ExcelFile(active_file)
    perf_list = []
    for sheet in xls.sheet_names:
        if "歷史樣本" in sheet:
            try:
                df_block = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=3, usecols="E:H")
                plan_d = pd.to_numeric(df_block.iloc[1, 2], errors='coerce')
                act_d = pd.to_numeric(df_block.iloc[1, 3], errors='coerce')
                if plan_d > 0 and act_d > 0:
                    case_ratio = act_d / plan_d
                    for r in [1, 2]:
                        e_name = str(df_block.iloc[r, 0]).strip()
                        f_name = str(df_block.iloc[r, 1]).strip()
                        if e_name not in ["nan", "0", "", "None"]: perf_list.append({'name': e_name, 'type': 'EPC', 'ratio': case_ratio})
                        if f_name not in ["nan", "0", "", "None"]: perf_list.append({'name': f_name, 'type': 'PCM', 'ratio': case_ratio})
            except: continue
    df_perf = pd.DataFrame(perf_list)
    perf_summary = df_perf.groupby(['name', 'type'])['ratio'].mean().to_dict() if not df_perf.empty else {}

    target_case_name = st.sidebar.text_input("目標案件工作表", value="平均預測")
    
    if target_case_name in xls.sheet_names:
        df_init = pd.read_excel(xls, sheet_name=target_case_name, header=None, nrows=2, usecols="A:F")
        init_total_price = float(df_init.iloc[1, 4])
        
        # 參數設置
        st.sidebar.header("⚙️ 模擬參數")
        total_p = st.sidebar.number_input("總價金額", value=init_total_price)
        design_f = st.sidebar.number_input("設計金額", value=round(total_p * 0.02, 0))
        const_p = total_p - design_f
        contract_d = st.sidebar.date_input("合約起始日期")
        start_d = st.sidebar.date_input("預計開工日期", value=contract_d + timedelta(days=365))
        manual_dur = st.sidebar.number_input("基準施工總天數", value=1100)
        num_sims = st.sidebar.slider("蒙地卡羅次數", 100, 1000, 400)
        
        env_ratio = st.sidebar.slider("修正倍率", 0.5, 2.5, 1.0)
        use_protection = st.sidebar.toggle("啟動進度保護", value=True)

        # 模擬計算
        start_dt = pd.to_datetime(start_d)
        target_df, date_col, cum_col, _ = clean_and_process(pd.read_excel(xls, target_case_name), start_dt)
        
        case_list, case_info = [], {}
        for sheet in xls.sheet_names:
            if sheet == target_case_name or "歷史樣本" not in sheet: continue
            df_h, _, _, _ = clean_and_process(pd.read_excel(xls, sheet), None)
            if df_h is not None:
                df_h["案件"] = sheet; case_list.append(df_h)
                t_y = np.interp(np.linspace(0, 100, 100), target_df["天數_norm"].values, target_df["累計_norm"].ffill().fillna(0).values)
                c_y = np.interp(np.linspace(0, 100, 100), df_h["天數_norm"].values, df_h["累計_norm"].ffill().fillna(0).values)
                case_info[sheet] = {'similarity': max(0.001, np.corrcoef(t_y, c_y)[0, 1])}
        
        merged_df = pd.concat(case_list, ignore_index=True)
        top_names = [n for n, _ in sorted(case_info.items(), key=lambda x: x[1]['similarity'], reverse=True)[:4]]
        weights = np.array([case_info[n]['similarity'] for n in top_names])**2
        weights /= weights.sum()

        last_p = target_df["累計_norm"].iloc[-1]
        last_d = (target_df[date_col].iloc[-1] - start_dt).days if last_p > 0 else 0
        prog_steps = np.linspace(last_p, 100, 101)
        sim_matrix = []

        for _ in range(num_sims):
            curve_days, curr_d = [], last_d
            for l, h in zip(np.linspace(last_p, 90, 10), np.linspace(last_p+10, 100, 10)):
                case = np.random.choice(top_names, p=weights)
                sub = merged_df[(merged_df["案件"] == case) & (merged_df["累計_norm"] >= l) & (merged_df["累計_norm"] <= h)]
                if len(sub) >= 2:
                    y_s = sub["天數_norm"].to_numpy() / 100 * manual_dur
                    interp_d = np.interp(np.linspace(l, h, 20), sub["累計_norm"].to_numpy(), y_s)
                    inc = max(1, interp_d[-1] - interp_d[0])
                    curr_d += inc
                    curve_days.extend(np.linspace(curr_d-inc, curr_d, 20).tolist())
            if curve_days:
                sim_matrix.append(np.interp(prog_steps, np.linspace(last_p, 100, len(curve_days)), curve_days))

        sim_matrix = np.atleast_2d(sim_matrix)
        mean_c = np.nanmean(sim_matrix, axis=0)
        p10, p90 = np.nanpercentile(sim_matrix, 10, axis=0), np.nanpercentile(sim_matrix, 90, axis=0)
        p15, p85 = np.nanpercentile(sim_matrix, 15, axis=0), np.nanpercentile(sim_matrix, 85, axis=0)
        p25, p75 = np.nanpercentile(sim_matrix, 25, axis=0), np.nanpercentile(sim_matrix, 75, axis=0)

        # 繪圖
        def to_dates(curve): return [start_dt + timedelta(days=int(d * env_ratio)) for d in curve]
        u_days = (np.concatenate([target_df["天數"].values, mean_c])) * env_ratio
        u_prog = np.concatenate([target_df["累計_norm"].ffill().fillna(0).values, prog_steps])
        s_idx = np.argsort(u_days); u_days, u_prog = u_days[s_idx], u_prog[s_idx]
        hover_pay = [f"{int(const_p * (np.interp(d * env_ratio, u_days, u_prog) - np.interp(max(0, (d-30) * env_ratio), u_days, u_prog)) / 100):,} 元" for d in mean_c]

        fig = go.Figure()
        fig.add_trace(go.Scatter(x=to_dates(p10)+to_dates(p90)[::-1], y=prog_steps.tolist()+prog_steps[::-1].tolist(), fill='toself', fillcolor='rgba(149,165,166,0.1)', name='90% 信賴區間', hoverinfo='skip'))
        fig.add_trace(go.Scatter(x=to_dates(p15)+to_dates(p85)[::-1], y=prog_steps.tolist()+prog_steps[::-1].tolist(), fill='toself', fillcolor='rgba(241,196,15,0.15)', name='70% 信賴區間', hoverinfo='skip'))
        fig.add_trace(go.Scatter(x=to_dates(p25)+to_dates(p75)[::-1], y=prog_steps.tolist()+prog_steps[::-1].tolist(), fill='toself', fillcolor='rgba(46,134,193,0.2)', name='50% 信賴區間', hoverinfo='skip'))
        fig.add_trace(go.Scatter(x=to_dates(mean_c), y=prog_steps, mode='lines', name='預測進度 (Mean)', line=dict(color='#3498db', width=3.5, dash='dash'), customdata=hover_pay, hovertemplate="日期: %{x}<br>進度: %{y:.1f}%<br>支用: %{customdata}<extra></extra>"))
        fig.update_layout(title=f"<b>{target_case_name} S-Curve 進度預測</b>", hovermode="x unified", template="plotly_white", height=600, legend=dict(orientation="h", y=1.1))
        st.plotly_chart(fig, use_container_width=True)

    else: st.error("找不到工作表")
else: st.info("請上傳 Excel 檔案。")
