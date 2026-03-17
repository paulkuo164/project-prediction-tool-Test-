import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objs as go
from datetime import timedelta
import calendar
import warnings
import io
import os


st.set_page_config(page_title="工程金流預測儀表板", layout="wide")

# 在這裡加入您的 LOGO
st.logo("logo.png", link="https://www.hurc.org.tw/hurc/hpage")

# ===== 🤐 系統設定 =====
warnings.filterwarnings('ignore')
np.seterr(divide='ignore', invalid='ignore')

st.set_page_config(page_title="工程金流預測儀表板", layout="wide")
st.title("🏗️ 工程進度預測與金流預估")

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

def get_contract_year(target_date, contract_start_date):
    if pd.isna(target_date) or target_date is None: return "未知年度"
    t_dt = pd.to_datetime(target_date)
    s_dt = pd.to_datetime(contract_start_date)
    months_delta = (t_dt.year - s_dt.year) * 12 + (t_dt.month - s_dt.month)
    year_num = (months_delta // 12) + 1
    return f"第 {int(year_num)} 年度"

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
    df[cum_col] = pd.to_numeric(df[cum_col], errors="coerce").ffill().fillna(0)
    max_p = df[cum_col].max()
    df["累計_norm"] = (df[cum_col] / (max_p if max_p > 0 else 1)) * 100
    df["天數_norm"] = (df["天數"] / (df["天數"].max() if df["天數"].max() > 0 else 1)) * 100
    return df, date_col, cum_col, actual_start

# ===== 📂 2. 資料來源與廠商績效抓取 =====
st.sidebar.header("📂 資料來源")
default_file = "歷史案件資料(工期+金額)改.xlsx"
has_internal = os.path.exists(default_file)

if has_internal:
    data_option = st.sidebar.radio("數據源", ["使用內建歷史檔", "手動上傳新檔案"])
    active_file = default_file if data_option == "使用內建歷史檔" else st.sidebar.file_uploader("上傳 Excel", type=["xlsx"])
else:
    active_file = st.sidebar.file_uploader("上傳歷史案件 Excel", type=["xlsx"])

if active_file:
    xls = pd.ExcelFile(active_file)
    
    # --- 🏢 廠商績效自動抓取 (從各歷史樣本 E2:H3) ---
    perf_list = []
    for sheet in xls.sheet_names:
        if "歷史樣本" in sheet:
            try:
                # 讀取 E2:H3 (header=None, row 1~2, col 4~7)
                df_block = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=3, usecols="E:H").fillna(0)
                plan_d = df_block.iloc[1, 2] # G2
                act_d = df_block.iloc[1, 3]  # H2
                case_ratio = act_d / plan_d if (isinstance(plan_d, (int, float)) and plan_d > 0) else 1.0
                
                # E2, E3 (EPC); F2, F3 (PCM)
                for r in [1, 2]:
                    e_name, f_name = df_block.iloc[r, 0], df_block.iloc[r, 1]
                    if e_name != 0 and str(e_name).strip() != "":
                        perf_list.append({'name': str(e_name).strip(), 'type': 'EPC', 'ratio': case_ratio})
                    if f_name != 0 and str(f_name).strip() != "":
                        perf_list.append({'name': str(f_name).strip(), 'type': 'PCM', 'ratio': case_ratio})
            except: continue
            
    df_perf = pd.DataFrame(perf_list)
    perf_summary = df_perf.groupby(['name', 'type'])['ratio'].mean().to_dict() if not df_perf.empty else {}

    target_case_name = st.sidebar.text_input("目標案件工作表", value="平均預測")
    
    if target_case_name in xls.sheet_names:
        df_init = pd.read_excel(xls, sheet_name=target_case_name, header=None, nrows=2, usecols="A:F")
        init_contract_date = pd.to_datetime(df_init.iloc[1, 0]).date()
        init_total_price = float(df_init.iloc[1, 4])
        
        # --- ⚙️ 模擬參數調整 ---
        st.sidebar.markdown("---")
        st.sidebar.header("⚙️ 模擬參數調整")
        total_p = st.sidebar.number_input("總價金額 (元)", value=init_total_price, step=1000000.0)
        design_f = st.sidebar.number_input("統包設計金額", value=round(total_p * 0.02, 0), step=10000.0)
        const_p = total_p - design_f
        contract_d = st.sidebar.date_input("合約起始日期", value=init_contract_date)
        start_d = st.sidebar.date_input("預計開工日期", value=contract_d + timedelta(days=365))
        manual_dur = st.sidebar.number_input("基準預期施工總天數", value=1100)
        num_sims = st.sidebar.slider("蒙地卡羅模擬次數", 100, 1000, 400)

        # --- ✨ 廠商風險修正連動 ---
        st.sidebar.markdown("---")
        st.sidebar.subheader("⚖️ 風險修正係數 (廠商)")
        all_e = sorted(list(set([k[0] for k in perf_summary.keys() if k[1] == 'EPC'])))
        all_f = sorted(list(set([k[0] for k in perf_summary.keys() if k[1] == 'PCM'])))
        sel_e = st.sidebar.multiselect("統包商 (E2/E3)", options=all_e)
        sel_f = st.sidebar.multiselect("監造單位 (F2/F3)", options=all_f)
        
        r_vals = [perf_summary.get((e, 'EPC'), 1.0) for e in sel_e] + [perf_summary.get((f, 'PCM'), 1.0) for f in sel_f]
        vendor_suggested = np.mean(r_vals) if r_vals else 1.0
        
        use_env_adj = st.sidebar.toggle("啟用風險修正係數", value=True if r_vals else False)
        env_ratio = st.sidebar.slider("修正倍率", 0.5, 2.0, float(vendor_suggested)) if use_env_adj else 1.0
        use_protection = st.sidebar.toggle("啟動進度保護機制", value=True)

        # --- 核心數據處理與模擬 ---
        start_dt = pd.to_datetime(start_d)
        target_df, date_col, cum_col, _ = clean_and_process(pd.read_excel(xls, target_case_name), start_dt)
        
        case_list, case_info = [], {}
        for sheet in xls.sheet_names:
            if sheet == target_case_name or "歷史樣本" not in sheet: continue
            df_h, _, _, _ = clean_and_process(pd.read_excel(xls, sheet), None)
            if df_h is not None:
                df_h["案件"] = sheet; case_list.append(df_h)
                x_i = np.linspace(0, 100, 100)
                t_y = np.interp(x_i, target_df["天數_norm"].values, target_df["累計_norm"].values)
                c_y = np.interp(x_i, df_h["天數_norm"].values, df_h["累計_norm"].values)
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
                    inc = max(1, interp_d[-1] - interp_d[0]); curr_d += inc
                    curve_days.extend(np.linspace(curr_d-inc, curr_d, 20).tolist())
            
            if use_protection:
                if curve_days: sim_matrix.append(np.interp(prog_steps, np.linspace(last_p, 100, len(curve_days)), curve_days))
            else:
                sim_matrix.append(np.interp(prog_steps, np.linspace(last_p, 100, len(curve_days)), curve_days) if curve_days else np.zeros_like(prog_steps))

        sim_matrix = np.atleast_2d(sim_matrix)
        mean_c = np.nanmean(sim_matrix, axis=0)
        p10, p90 = np.nanpercentile(sim_matrix, 10, axis=0), np.nanpercentile(sim_matrix, 90, axis=0)
        p15, p85 = np.nanpercentile(sim_matrix, 15, axis=0), np.nanpercentile(sim_matrix, 85, axis=0)
        p25, p75 = np.nanpercentile(sim_matrix, 25, axis=0), np.nanpercentile(sim_matrix, 75, axis=0)

# --- 📈 3. 圖表渲染 (動態面積視覺化版) ---
        fig = go.Figure()

        # 1. 加入信賴區間 (背景層)
        fig.add_trace(go.Scatter(x=to_dates(p10)+to_dates(p90)[::-1], y=prog_steps.tolist()+prog_steps[::-1].tolist(), fill='toself', fillcolor='rgba(149,165,166,0.1)', line=dict(color='rgba(255,255,255,0)'), name='90% 信賴區間', hoverinfo='skip'))
        fig.add_trace(go.Scatter(x=to_dates(p15)+to_dates(p85)[::-1], y=prog_steps.tolist()+prog_steps[::-1].tolist(), fill='toself', fillcolor='rgba(241,196,15,0.15)', line=dict(color='rgba(255,255,255,0)'), name='70% 信賴區間', hoverinfo='skip'))
        fig.add_trace(go.Scatter(x=to_dates(p25)+to_dates(p75)[::-1], y=prog_steps.tolist()+prog_steps[::-1].tolist(), fill='toself', fillcolor='rgba(46,134,193,0.2)', line=dict(color='rgba(255,255,255,0)'), name='50% 信賴區間', hoverinfo='skip'))

        # 2. 主預測線
        fig.add_trace(go.Scatter(
            x=to_dates(mean_c), 
            y=prog_steps, 
            mode='lines', 
            name='預測進度 (Mean)', 
            line=dict(color='#3498db', width=3.5, dash='dash'),
            customdata=hover_custom_data,
            hovertemplate=(
                "<b>📅 日期</b>: %{x|%Y-%m-%d}<br><br>" +
                "<b>📈 進度驗證：</b><br>" +
                "└ 樂觀 (P10): %{customdata[2]}<br>" +
                "└ 平均 (Mean): %{customdata[3]}<br>" +
                "└ 悲觀 (P90): %{customdata[4]}<br><br>" +
                "<b>💰 金流預估：</b><br>" +
                "└ 預估當期支付: %{customdata[0]}<br>" +
                "└ 樂觀/悲觀價差: <span style='color:red'>%{customdata[1]}</span>" +
                "<extra></extra>"
            )
        ))

        # 3. ✨ 視覺化核心：設定 Layout 讓滑鼠移過時顯示垂直區塊 (Spikeline 擴展)
        fig.update_layout(
            title=f"<b>{target_case_name} S-Curve 進度預測與風險分析</b>",
            hovermode="x unified", # 這是關鍵，讓滑鼠在 X 軸移動時統一觸發
            template="plotly_white",
            height=650,
            xaxis=dict(
                showspikes=True, # 顯示垂直引導線
                spikemode="across",
                spikesnap="cursor",
                spikethickness=40, # 這裡調粗線條，模擬「面積區塊」的效果
                spikecolor="rgba(52, 152, 219, 0.15)", # 使用極淡的藍色區塊
            ),
            yaxis=dict(range=[0, 105]),
            legend=dict(orientation="h", y=1.1)
        )

        st.plotly_chart(fig, use_container_width=True)

        # === 💰 第二部分：全週期金流分析 ===
        st.markdown("---")
        st.subheader("💰 全週期金流分析與互動排程")

        if 'design_df' not in st.session_state:
            st.session_state.design_df = pd.DataFrame([
                {"期別": "設計一期", "基準點": "合約起始", "相對月數": 3, "比例": 0.10},
                {"期別": "設計二期", "基準點": "合約起始", "相對月數": 6, "比例": 0.15},
                {"期別": "設計三期", "基準點": "合約起始", "相對月數": 9, "比例": 0.20},
                {"期別": "設計四期", "基準點": "預計開工", "相對月數": 6, "比例": 0.45},
                {"期別": "設計五期", "基準點": "預計完工", "相對月數": 1, "比例": 0.10},
            ])

        with st.expander("🛠️ 調整設計款支付時程", expanded=True):
            edited_design_df = st.data_editor(
                st.session_state.design_df,
                column_config={
                    "期別": st.column_config.TextColumn("款項名稱"),
                    "基準點": st.column_config.SelectboxColumn("日期基準", options=["合約起始", "預計開工", "預計完工"]),
                    "相對月數": st.column_config.NumberColumn("延後月數", min_value=0, max_value=120, step=1),
                    "比例": st.column_config.NumberColumn("支付比例", min_value=0.0, max_value=1.0, format="%.2f")
                },
                num_rows="dynamic", use_container_width=True, key="design_editor_integrated"
            )
            st.session_state.design_df = edited_design_df

        mean_finish_dt = start_dt + timedelta(days=int(mean_c[-1] * env_ratio))
        pay_data = []
        for _, row in edited_design_df.iterrows():
            base_ref = pd.to_datetime(contract_d) if row["基準點"] == "合約起始" else (pd.to_datetime(start_d) if row["基準點"] == "預計開工" else mean_finish_dt)
            p_date = get_payment_date(get_month_end(base_ref + pd.DateOffset(months=int(row["相對月數"]))))
            pay_data.append({"期別": row["期別"], "性質": "設計款", "支付日": p_date, "金額": int(design_f * row["比例"])})

        curr_m = pd.to_datetime(start_d).replace(day=1)
        prev_p = 0
        while curr_m <= (mean_finish_dt + timedelta(days=90)): 
            m_end = get_month_end(curr_m)
            if m_end >= start_dt:
                ref_day = (m_end - start_dt).days / env_ratio
                cp = np.interp(ref_day, u_days / env_ratio, u_prog)
                if cp > prev_p:
                    amt = int(const_p * (cp - prev_p) / 100)
                    if amt > 0:
                        pay_data.append({"期別": f"工程估驗 {m_end.strftime('%Y/%m')}", "性質": "工程款", "支付日": get_payment_date(m_end), "金額": amt})
                    prev_p = cp
            if prev_p >= 100: break
            curr_m += pd.DateOffset(months=1)

        df_pay = pd.DataFrame(pay_data)
        df_pay['月份'] = df_pay['支付日'].dt.strftime('%Y-%m')
        df_monthly = df_pay.groupby('月份')['金額'].sum().reset_index().sort_values('月份')

        # --- 樣式 TAB 切換 ---
        tab1, tab2, tab3 = st.tabs(["📊 每月支出趨勢", "📜 詳細金流明細", "📅 工期情境總結"])

        with tab1:
            fig_bar = go.Figure(data=[go.Bar(
                x=df_monthly['月份'], y=df_monthly['金額'], 
                marker_color='#2ecc71',
                text=[f"{v/10000:,.0f}萬" for v in df_monthly['金額']], 
                textposition='outside'
            )])
            fig_bar.update_layout(title="<b>各月預計付款金額 (元)</b>", template="plotly_white", height=450)
            st.plotly_chart(fig_bar, use_container_width=True)

        with tab2:
            show_df = df_pay.sort_values("支付日").copy()
            show_df["合約年度"] = show_df["支付日"].apply(lambda x: get_contract_year(x, contract_d))
            for year in show_df["合約年度"].unique():
                with st.expander(f"📅 {year} 明細", expanded=True):
                    df_y = show_df[show_df["合約年度"] == year].copy()
                    df_y["金額(元)"] = df_y["金額"].apply(lambda x: f"{x:,}")
                    st.table(df_y[["支付日", "期別", "性質", "金額(元)"]])
                    st.markdown(f"**💰 {year} 撥款總計： `{df_y['金額'].sum():,}` 元**")

        with tab3:
            col1, col2 = st.columns(2)
            with col1:
                st.write("#### 🗓️ 關鍵里程碑預估")
                st.table(pd.DataFrame([
                    {"情境": "樂觀 (P10)", "預計完工日期": (start_dt + timedelta(days=int(p10[-1] * env_ratio))).date()},
                    {"情境": "平均 (Mean)", "預計完工日期": (start_dt + timedelta(days=int(mean_c[-1] * env_ratio))).date()},
                    {"情境": "悲觀 (P90)", "預計完工日期": (start_dt + timedelta(days=int(p90[-1] * env_ratio))).date()}
                ]))
            with col2:
                st.write("#### 📥 報表下載")
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_pay.to_excel(writer, sheet_name='金流分析', index=False)
                st.download_button("📥 下載預測報表 (.xlsx)", data=buffer.getvalue(), file_name=f"{target_case_name}_預測.xlsx")

    else:
        st.error(f"找不到工作表「{target_case_name}」")
else:
    st.info("💡 請上傳檔案以開始分析。系統將自動從歷史樣本提取廠商績效。")
