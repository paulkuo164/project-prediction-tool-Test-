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
                df_block = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=3, usecols="E:H").fillna(0)
                plan_d = df_block.iloc[1, 2]
                act_d = df_block.iloc[1, 3]
                case_ratio = act_d / plan_d if (isinstance(plan_d, (int, float)) and plan_d > 0) else 1.0
                
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
        total_p = st.sidebar.number_input("決標金額 (元)", value=init_total_price, step=1000000.0,
                                          help="本案決標總金額，為所有費用拆分的基礎")
        
        # === 💼 費用結構拆分 ===
        st.sidebar.markdown("---")
        st.sidebar.subheader("💼 費用結構拆分")
        
        # 設計費 (決標金額的百分比)
        design_pct = st.sidebar.slider("設計費比例 (%)", 0.0, 10.0, 3.0, 0.1,
                                       help="設計費佔決標金額之比例，預設 3%")
        design_f = round(total_p * design_pct / 100, 0)
        st.sidebar.caption(f"　設計費：**{design_f:,.0f} 元**")
        
        # 統包施工費 = 決標金額 - 設計費 (所有後續費用的基數)
        const_p = total_p - design_f
        st.sidebar.markdown(f"📌 **統包施工費：`{const_p:,.0f}` 元**")
        st.sidebar.caption("（= 決標金額 − 設計費，下列各項費用之計算基數）")
        
        # 專管及監造服務費 (可調)
        pcm_pct = st.sidebar.slider("專管及監造服務費 (%)", 0.0, 10.0, 4.0, 0.1,
                                    help="包含耐震特別監督，預設 4%，可調")
        pcm_fee = round(const_p * pcm_pct / 100, 0)
        st.sidebar.caption(f"　專管監造費：**{pcm_fee:,.0f} 元**")
        
        # 準備金 (鎖定 2%)
        reserve_fee = round(const_p * 0.02, 0)
        st.sidebar.caption(f"🔒 準備金 (2%)：**{reserve_fee:,.0f} 元**")
        
        # 物調款 (可調)
        price_adj_pct = st.sidebar.slider("物調款 (%)", 0.0, 20.0, 8.0, 0.1,
                                          help="物價調整款，預設 8%，可調")
        price_adj_fee = round(const_p * price_adj_pct / 100, 0)
        st.sidebar.caption(f"　物調款：**{price_adj_fee:,.0f} 元**")
        
        # 外管補助費 (鎖定 1%)
        external_fee = round(const_p * 0.01, 0)
        st.sidebar.caption(f"🔒 外管補助費 (1%)：**{external_fee:,.0f} 元**")
        
        # 公共藝術 (鎖定 1%)
        public_art_fee = round(const_p * 0.01, 0)
        st.sidebar.caption(f"🔒 公共藝術 (1%)：**{public_art_fee:,.0f} 元**")
        
        # 其他費用 (手動輸入)
        # 建議值：統包施工費 × 15% − 設計費 − 專管 − 準備金 − 物調款
        suggested_other = max(0, round(const_p * 0.15 - design_f - pcm_fee - reserve_fee - price_adj_fee, 0))
        other_fee = st.sidebar.number_input("其他費用 (元)", value=float(suggested_other), step=10000.0,
                                            help=f"建議值 {suggested_other:,.0f} 元（= 統包施工費×15% − 設計費 − 專管 − 準備金 − 物調款）")
        
        # 費用總表
        total_all_fees = design_f + pcm_fee + reserve_fee + price_adj_fee + external_fee + public_art_fee + other_fee
        st.sidebar.markdown(f"💰 **費用總計：`{total_all_fees:,.0f}` 元**")
        
        st.sidebar.markdown("---")
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

        # === 📋 主畫面：費用拆分摘要 ===
        with st.expander("💼 本案費用結構拆分（點擊展開/收合）", expanded=False):
            fee_breakdown_df = pd.DataFrame([
                {"項目": "設計費",               "比例(%)": f"{design_pct:.2f}",   "金額(元)": f"{design_f:,.0f}",       "基數": "決標金額",   "備註": "可調"},
                {"項目": "統包施工費",           "比例(%)": "—",                   "金額(元)": f"{const_p:,.0f}",         "基數": "—",         "備註": "決標金額 − 設計費"},
                {"項目": "專管及監造服務費",     "比例(%)": f"{pcm_pct:.2f}",      "金額(元)": f"{pcm_fee:,.0f}",         "基數": "統包施工費", "備註": "含耐震特別監督，可調"},
                {"項目": "準備金",               "比例(%)": "2.00",                "金額(元)": f"{reserve_fee:,.0f}",     "基數": "統包施工費", "備註": "🔒 鎖定"},
                {"項目": "物調款",               "比例(%)": f"{price_adj_pct:.2f}","金額(元)": f"{price_adj_fee:,.0f}",   "基數": "統包施工費", "備註": "可調"},
                {"項目": "外管補助費",           "比例(%)": "1.00",                "金額(元)": f"{external_fee:,.0f}",    "基數": "統包施工費", "備註": "🔒 鎖定"},
                {"項目": "公共藝術",             "比例(%)": "1.00",                "金額(元)": f"{public_art_fee:,.0f}",  "基數": "統包施工費", "備註": "🔒 鎖定"},
                {"項目": "其他費用",             "比例(%)": "—",                   "金額(元)": f"{other_fee:,.0f}",       "基數": "—",         "備註": "手動輸入"},
            ])
            st.table(fee_breakdown_df)
            c1, c2, c3 = st.columns(3)
            c1.metric("決標金額", f"{total_p:,.0f} 元")
            c2.metric("統包施工費", f"{const_p:,.0f} 元")
            c3.metric("各項費用合計", f"{total_all_fees:,.0f} 元")

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

        # --- 📈 3. 圖表渲染 ---
        def to_dates(curve): return [start_dt + timedelta(days=int(d * env_ratio)) for d in curve]
        
        u_days = (np.concatenate([target_df["天數"].values, mean_c])) * env_ratio
        u_prog = np.concatenate([target_df["累計_norm"].values, prog_steps])
        s_idx = np.argsort(u_days); u_days, u_prog = u_days[s_idx], u_prog[s_idx]

        # === Hover 資料計算 ===
        # 核心原則：當期(30天) vs 累計 分開算，且三條情境線各自回推自己的 30 天前進度
        # 這樣每個數字的時間尺度都一致、可直接對照
        PERIOD_DAYS = 30  # 當期 = 近 30 天（對齊下方每月付款表）

        hover_custom_data = []
        for i, d_mean in enumerate(mean_c):
            # 目前這一點的「進度」（y 軸值）—— 三條線都指向同一個進度點
            prog_now = prog_steps[i]
            
            # 三種情境在「達到此進度」所需的天數
            d_p10 = p10[i]
            d_p90 = p90[i]
            # d_mean 本身就是 mean 情境下的天數
            
            # --- 🧮 累計金額 (到達此進度時，各情境已完成的金額) ---
            # 注意：三條情境線在同一個 prog_now 下的累計「工程款」其實相同
            # 差異來自「同一天」各情境進度不同 → 用 hover 日期 x 軸(mean 日期)當基準
            cum_mean = const_p * prog_now / 100
            # 若要問「在 mean 的這一天」，各情境分別累積到多少：
            prog_p10_at_d = np.interp(d_mean, p10, prog_steps)  # mean 的這一天，樂觀已做到幾%
            prog_p90_at_d = np.interp(d_mean, p90, prog_steps)  # mean 的這一天，悲觀已做到幾%
            cum_p10 = const_p * prog_p10_at_d / 100
            cum_p90 = const_p * prog_p90_at_d / 100
            cum_gap = cum_p10 - cum_p90  # 正值代表樂觀領先
            
            # --- 📆 當期金額 (近 30 天，各情境各自計算) ---
            # 每條線「這一天」的進度 vs 「30 天前」的進度差 → 乘以工程款基數
            # 樂觀、悲觀各自用自己的曲線插值，不共用 mean 的時間
            prog_mean_prev = np.interp(max(0, d_mean - PERIOD_DAYS), mean_c, prog_steps)
            prog_p10_prev  = np.interp(max(0, d_p10  - PERIOD_DAYS), p10,    prog_steps)
            prog_p90_prev  = np.interp(max(0, d_p90  - PERIOD_DAYS), p90,    prog_steps)
            
            pay_mean = const_p * (prog_now - prog_mean_prev) / 100
            pay_p10  = const_p * (prog_now - prog_p10_prev)  / 100
            pay_p90  = const_p * (prog_now - prog_p90_prev)  / 100
            period_gap = pay_p10 - pay_p90  # 正值代表樂觀當期支付較多
            
            hover_custom_data.append([
                # [0-2] 進度驗證
                f"{prog_p10_at_d:.2f}%",
                f"{prog_now:.2f}%",
                f"{prog_p90_at_d:.2f}%",
                # [3-6] 當期金流（近 30 天）
                f"{int(pay_p10):,} 元",
                f"{int(pay_mean):,} 元",
                f"{int(pay_p90):,} 元",
                f"{int(period_gap):,} 元",
                # [7-10] 累計金流（至本日）
                f"{int(cum_p10):,} 元",
                f"{int(cum_mean):,} 元",
                f"{int(cum_p90):,} 元",
                f"{int(cum_gap):,} 元",
            ])

        fig = go.Figure()
        fig.add_trace(go.Scatter(x=to_dates(p10)+to_dates(p90)[::-1], y=prog_steps.tolist()+prog_steps[::-1].tolist(), fill='toself', fillcolor='rgba(149,165,166,0.1)', line=dict(color='rgba(255,255,255,0)'), name='90% 信賴區間', hoverinfo='skip'))
        fig.add_trace(go.Scatter(x=to_dates(p15)+to_dates(p85)[::-1], y=prog_steps.tolist()+prog_steps[::-1].tolist(), fill='toself', fillcolor='rgba(241,196,15,0.15)', line=dict(color='rgba(255,255,255,0)'), name='70% 信賴區間', hoverinfo='skip'))
        fig.add_trace(go.Scatter(x=to_dates(p25)+to_dates(p75)[::-1], y=prog_steps.tolist()+prog_steps[::-1].tolist(), fill='toself', fillcolor='rgba(46,134,193,0.2)', line=dict(color='rgba(255,255,255,0)'), name='50% 信賴區間', hoverinfo='skip'))
        
        fig.add_trace(go.Scatter(
            x=to_dates(mean_c), 
            y=prog_steps, 
            mode='lines', 
            name='預測進度 (Mean)', 
            line=dict(color='#3498db', width=3.5, dash='dash'),
            customdata=hover_custom_data,
            hovertemplate=(
                "<b>📅 日期</b>: %{x|%Y-%m-%d}<br>" +
                "<span style='color:#888'>(x 軸為 Mean 情境日期)</span><br><br>" +
                "<b>📈 進度驗證 (本日 %)</b><br>" +
                "└ 樂觀 (P10): %{customdata[0]}<br>" +
                "└ 平均 (Mean): %{customdata[1]}<br>" +
                "└ 悲觀 (P90): %{customdata[2]}<br><br>" +
                "<b>📆 當期金流 (近 30 天)</b><br>" +
                "└ 樂觀當期: %{customdata[3]}<br>" +
                "└ 平均當期: %{customdata[4]}<br>" +
                "└ 悲觀當期: %{customdata[5]}<br>" +
                "└ <b>當期風險價差</b>: <span style='color:#e74c3c'>%{customdata[6]}</span><br><br>" +
                "<b>📊 累計金流 (至本日)</b><br>" +
                "└ 樂觀累計: %{customdata[7]}<br>" +
                "└ 平均累計: %{customdata[8]}<br>" +
                "└ 悲觀累計: %{customdata[9]}<br>" +
                "└ <b>累計風險價差</b>: <span style='color:#e74c3c'>%{customdata[10]}</span>" +
                "<extra></extra>"
            )
        ))
        
        fig.update_layout(
            title=f"<b>{target_case_name} S-Curve 進度預測與風險分析</b>",
            hovermode="x unified",
            template="plotly_white",
            height=650,
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
                # === 匯出多工作表 Excel ===
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    # Sheet 1: 金流分析
                    df_pay.to_excel(writer, sheet_name='金流分析', index=False)
                    
                    # Sheet 2: 費用拆分
                    fee_export_df = pd.DataFrame([
                        {"項目": "決標金額",             "比例(%)": "",              "金額(元)": total_p,         "基數": "",           "備註": "本案決標總金額"},
                        {"項目": "設計費",               "比例(%)": design_pct,      "金額(元)": design_f,        "基數": "決標金額",   "備註": "可調"},
                        {"項目": "統包施工費",           "比例(%)": "",              "金額(元)": const_p,         "基數": "",           "備註": "決標金額 − 設計費"},
                        {"項目": "專管及監造服務費",     "比例(%)": pcm_pct,         "金額(元)": pcm_fee,         "基數": "統包施工費", "備註": "含耐震特別監督，可調"},
                        {"項目": "準備金",               "比例(%)": 2.00,            "金額(元)": reserve_fee,     "基數": "統包施工費", "備註": "鎖定"},
                        {"項目": "物調款",               "比例(%)": price_adj_pct,   "金額(元)": price_adj_fee,   "基數": "統包施工費", "備註": "可調"},
                        {"項目": "外管補助費",           "比例(%)": 1.00,            "金額(元)": external_fee,    "基數": "統包施工費", "備註": "鎖定"},
                        {"項目": "公共藝術",             "比例(%)": 1.00,            "金額(元)": public_art_fee,  "基數": "統包施工費", "備註": "鎖定"},
                        {"項目": "其他費用",             "比例(%)": "",              "金額(元)": other_fee,       "基數": "",           "備註": "手動輸入"},
                        {"項目": "─────────",           "比例(%)": "",              "金額(元)": "",              "基數": "",           "備註": ""},
                        {"項目": "各項費用合計",         "比例(%)": "",              "金額(元)": total_all_fees,  "基數": "",           "備註": "設計費+專管+準備金+物調+外管+公共藝術+其他"},
                    ])
                    fee_export_df.to_excel(writer, sheet_name='費用拆分', index=False)
                    
                    # 設定欄寬
                    workbook = writer.book
                    money_fmt = workbook.add_format({'num_format': '#,##0'})
                    pct_fmt = workbook.add_format({'num_format': '0.00'})
                    
                    ws_fee = writer.sheets['費用拆分']
                    ws_fee.set_column('A:A', 22)
                    ws_fee.set_column('B:B', 10, pct_fmt)
                    ws_fee.set_column('C:C', 18, money_fmt)
                    ws_fee.set_column('D:D', 14)
                    ws_fee.set_column('E:E', 30)
                    
                    ws_pay = writer.sheets['金流分析']
                    ws_pay.set_column('A:A', 22)
                    ws_pay.set_column('B:B', 10)
                    ws_pay.set_column('C:C', 14)
                    ws_pay.set_column('D:D', 16, money_fmt)
                    ws_pay.set_column('E:E', 10)
                
                st.download_button("📥 下載預測報表 (.xlsx)", data=buffer.getvalue(), 
                                   file_name=f"{target_case_name}_預測.xlsx")

    else:
        st.error(f"找不到工作表「{target_case_name}」")
else:
    st.info("💡 請上傳檔案以開始分析。系統將自動從歷史樣本提取廠商績效。")
