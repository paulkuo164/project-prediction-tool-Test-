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

        design_pct = st.sidebar.slider("設計費比例 (%)", 0.0, 10.0, 3.0, 0.1,
                                       help="設計費佔決標金額之比例，預設 3%")
        design_f = round(total_p * design_pct / 100, 0)
        st.sidebar.caption(f"　設計費：**{design_f:,.0f} 元**")

        const_p = total_p - design_f
        st.sidebar.markdown(f"📌 **統包施工費：`{const_p:,.0f}` 元**")
        st.sidebar.caption("（= 決標金額 − 設計費，下列各項費用之計算基數）")

        pcm_pct = st.sidebar.slider("專管及監造服務費 (%)", 0.0, 10.0, 4.0, 0.1,
                                    help="包含耐震特別監督，預設 4%，可調")
        pcm_fee = round(const_p * pcm_pct / 100, 0)
        st.sidebar.caption(f"　專管監造費：**{pcm_fee:,.0f} 元**")

        reserve_fee = round(const_p * 0.02, 0)
        st.sidebar.caption(f"🔒 準備金 (2%)：**{reserve_fee:,.0f} 元**")

        price_adj_pct = st.sidebar.slider("物調款 (%)", 0.0, 20.0, 8.0, 0.1,
                                          help="物價調整款，預設 8%，可調")
        price_adj_fee = round(const_p * price_adj_pct / 100, 0)
        st.sidebar.caption(f"　物調款：**{price_adj_fee:,.0f} 元**")

        external_fee = round(const_p * 0.01, 0)
        st.sidebar.caption(f"🔒 外管補助費 (1%)：**{external_fee:,.0f} 元**")

        public_art_fee = round(const_p * 0.01, 0)
        st.sidebar.caption(f"🔒 公共藝術 (1%)：**{public_art_fee:,.0f} 元**")

        suggested_other = max(0, round(const_p * 0.15 - design_f - pcm_fee - reserve_fee - price_adj_fee, 0))
        other_fee = st.sidebar.number_input("其他費用 (元)", value=float(suggested_other), step=10000.0,
                                            help=f"建議值 {suggested_other:,.0f} 元（= 統包施工費×15% − 設計費 − 專管 − 準備金 − 物調款）")

        total_all_fees = design_f + pcm_fee + reserve_fee + price_adj_fee + external_fee + public_art_fee + other_fee
        st.sidebar.markdown(f"💰 **費用總計：`{total_all_fees:,.0f}` 元**")

        st.sidebar.markdown("---")
        contract_d = st.sidebar.date_input("合約起始日期", value=init_contract_date)
        start_d = st.sidebar.date_input("預計開工日期", value=contract_d + timedelta(days=365))
        manual_dur = st.sidebar.number_input("基準預期施工總天數", value=1100)
        num_sims = st.sidebar.slider("蒙地卡羅模擬次數", 100, 1000, 400)

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

        with st.expander("💼 本案費用結構拆分（點擊展開/收合）", expanded=False):
            fee_breakdown_df = pd.DataFrame([
                {"項目": "設計費",               "比例(%)": f"{design_pct:.2f}",   "預估金額(元)": f"{design_f:,.0f}",       "基數": "決標金額",   "備註": "可調"},
                {"項目": "統包施工費",           "比例(%)": "—",                   "預估金額(元)": f"{const_p:,.0f}",         "基數": "—",         "備註": "決標金額 − 設計費"},
                {"項目": "專管及監造服務費",     "比例(%)": f"{pcm_pct:.2f}",      "預估金額(元)": f"{pcm_fee:,.0f}",         "基數": "統包施工費", "備註": "含耐震特別監督，可調"},
                {"項目": "準備金",               "比例(%)": "2.00",                "預估金額(元)": f"{reserve_fee:,.0f}",     "基數": "統包施工費", "備註": "🔒 鎖定（不參與 S-curve 撥款）"},
                {"項目": "物調款",               "比例(%)": f"{price_adj_pct:.2f}","預估金額(元)": f"{price_adj_fee:,.0f}",   "基數": "統包施工費", "備註": "可調"},
                {"項目": "外管補助費",           "比例(%)": "1.00",                "預估金額(元)": f"{external_fee:,.0f}",    "基數": "統包施工費", "備註": "🔒 鎖定"},
                {"項目": "公共藝術",             "比例(%)": "1.00",                "預估金額(元)": f"{public_art_fee:,.0f}",  "基數": "統包施工費", "備註": "🔒 鎖定"},
                {"項目": "其他費用",             "比例(%)": "—",                   "預估金額(元)": f"{other_fee:,.0f}",       "基數": "—",         "備註": "手動輸入"},
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

        PERIOD_DAYS = 30

        hover_custom_data = []
        for i, d_mean in enumerate(mean_c):
            prog_now = prog_steps[i]
            d_p10 = p10[i]
            d_p90 = p90[i]

            cum_mean = const_p * prog_now / 100
            prog_p10_at_d = np.interp(d_mean, p10, prog_steps)
            prog_p90_at_d = np.interp(d_mean, p90, prog_steps)
            cum_p10 = const_p * prog_p10_at_d / 100
            cum_p90 = const_p * prog_p90_at_d / 100
            cum_gap = cum_p10 - cum_p90

            prog_mean_prev = np.interp(max(0, d_mean - PERIOD_DAYS), mean_c, prog_steps)
            prog_p10_prev  = np.interp(max(0, d_p10  - PERIOD_DAYS), p10,    prog_steps)
            prog_p90_prev  = np.interp(max(0, d_p90  - PERIOD_DAYS), p90,    prog_steps)

            pay_mean = const_p * (prog_now - prog_mean_prev) / 100
            pay_p10  = const_p * (prog_now - prog_p10_prev)  / 100
            pay_p90  = const_p * (prog_now - prog_p90_prev)  / 100
            period_gap = pay_p10 - pay_p90

            hover_custom_data.append([
                f"{prog_p10_at_d:.2f}%",
                f"{prog_now:.2f}%",
                f"{prog_p90_at_d:.2f}%",
                f"{int(pay_p10):,} 元",
                f"{int(pay_mean):,} 元",
                f"{int(pay_p90):,} 元",
                f"{int(period_gap):,} 元",
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
        if 'external_df' not in st.session_state:
            st.session_state.external_df = pd.DataFrame([
                {"期別": "外管補助", "基準點": "預計完工", "相對月數": 0, "比例": 1.00},
            ])
        if 'publicart_df' not in st.session_state:
            st.session_state.publicart_df = pd.DataFrame([
                {"期別": "公共藝術", "基準點": "預計完工", "相對月數": 0, "比例": 1.00},
            ])
        if 'other_df' not in st.session_state:
            st.session_state.other_df = pd.DataFrame([
                {"期別": "其他費用", "基準點": "預計完工", "相對月數": 0, "比例": 1.00},
            ])

        editor_config = {
            "期別":   st.column_config.TextColumn("款項名稱"),
            "基準點": st.column_config.SelectboxColumn("日期基準", options=["合約起始", "預計開工", "預計完工"]),
            "相對月數": st.column_config.NumberColumn("延後月數", min_value=0, max_value=120, step=1),
            "比例":   st.column_config.NumberColumn("支付比例", min_value=0.0, max_value=1.0, format="%.2f"),
        }

        with st.expander("🛠️ 調整自訂期別費用支付時程 (設計費 / 外管 / 公共藝術 / 其他)", expanded=True):
            ec1, ec2 = st.columns(2)
            with ec1:
                st.markdown(f"**✏️ 設計費**　總額 `{design_f:,.0f}` 元")
                st.session_state.design_df = st.data_editor(
                    st.session_state.design_df, column_config=editor_config,
                    num_rows="dynamic", use_container_width=True, key="design_editor")
                st.markdown(f"**🏗️ 外管補助費**　總額 `{external_fee:,.0f}` 元")
                st.session_state.external_df = st.data_editor(
                    st.session_state.external_df, column_config=editor_config,
                    num_rows="dynamic", use_container_width=True, key="external_editor")
            with ec2:
                st.markdown(f"**🎨 公共藝術**　總額 `{public_art_fee:,.0f}` 元")
                st.session_state.publicart_df = st.data_editor(
                    st.session_state.publicart_df, column_config=editor_config,
                    num_rows="dynamic", use_container_width=True, key="publicart_editor")
                st.markdown(f"**📌 其他費用**　總額 `{other_fee:,.0f}` 元")
                st.session_state.other_df = st.data_editor(
                    st.session_state.other_df, column_config=editor_config,
                    num_rows="dynamic", use_container_width=True, key="other_editor")

        mean_finish_dt = start_dt + timedelta(days=int(mean_c[-1] * env_ratio))

        pay_data = []
        for edited_df, total_amt, pay_kind in [
            (st.session_state.design_df,    design_f,       "設計費"),
            (st.session_state.external_df,  external_fee,   "外管補助"),
            (st.session_state.publicart_df, public_art_fee, "公共藝術"),
            (st.session_state.other_df,     other_fee,      "其他費用"),
        ]:
            for _, row in edited_df.iterrows():
                base_ref = (pd.to_datetime(contract_d) if row["基準點"] == "合約起始"
                            else pd.to_datetime(start_d) if row["基準點"] == "預計開工"
                            else mean_finish_dt)
                p_date = get_payment_date(get_month_end(base_ref + pd.DateOffset(months=int(row["相對月數"]))))
                pay_data.append({
                    "期別": row["期別"], "性質": pay_kind,
                    "支付日": p_date, "金額": int(total_amt * row["比例"])
                })

        scurve_components = [
            ("施工費",    const_p),
            ("專管監造",  pcm_fee),
            ("物調款",    price_adj_fee),
        ]
        scurve_base_total = sum(v for _, v in scurve_components)

        monthly_scurve_rows = []
        curr_m = pd.to_datetime(start_d).replace(day=1)
        prev_p = 0
        while curr_m <= (mean_finish_dt + timedelta(days=90)):
            m_end = get_month_end(curr_m)
            if m_end >= start_dt:
                ref_day = (m_end - start_dt).days / env_ratio
                cp = np.interp(ref_day, u_days / env_ratio, u_prog)
                if cp > prev_p:
                    period_ratio = (cp - prev_p) / 100
                    period_pct = cp - prev_p
                    pay_date = get_payment_date(m_end)
                    month_label = m_end.strftime('%Y/%m')
                    for comp_name, comp_base in scurve_components:
                        amt = int(comp_base * period_ratio)
                        if amt > 0:
                            pay_data.append({
                                "期別": f"工程估驗 {month_label} - {comp_name}",
                                "性質": comp_name,
                                "支付日": pay_date,
                                "金額": amt,
                            })
                            monthly_scurve_rows.append({
                                "月份": pay_date.strftime('%Y-%m'),
                                "支付日": pay_date,
                                "類別": comp_name,
                                "金額": amt,
                                "當月進度": period_pct,
                            })
                    prev_p = cp
            if prev_p >= 100: break
            curr_m += pd.DateOffset(months=1)

        df_pay = pd.DataFrame(pay_data)
        df_pay['月份'] = df_pay['支付日'].dt.strftime('%Y-%m')
        df_monthly = df_pay.groupby('月份')['金額'].sum().reset_index().sort_values('月份')

        if monthly_scurve_rows:
            df_scurve_long = pd.DataFrame(monthly_scurve_rows)
            df_progress = df_scurve_long.groupby("月份")["當月進度"].first().reset_index()
            df_scurve_pivot = df_scurve_long.pivot_table(
                index="月份", columns="類別", values="金額", aggfunc='sum', fill_value=0
            ).reset_index()
            for col in ["施工費", "專管監造", "物調款"]:
                if col not in df_scurve_pivot.columns:
                    df_scurve_pivot[col] = 0
            df_scurve_pivot = df_scurve_pivot[["月份", "施工費", "專管監造", "物調款"]]
            df_scurve_pivot["當月合計"] = df_scurve_pivot[["施工費", "專管監造", "物調款"]].sum(axis=1)
            df_scurve_pivot = df_scurve_pivot.merge(df_progress, on="月份", how="left")
            df_scurve_pivot = df_scurve_pivot[["月份", "當月進度", "施工費", "專管監造", "物調款", "當月合計"]]
        else:
            df_scurve_pivot = pd.DataFrame(columns=["月份", "當月進度", "施工費", "專管監造", "物調款", "當月合計"])

        WIDE_COLS = ["期別", "設計費", "外管補助", "公共藝術", "其他費用",
                     "專管監造", "物調款", "施工費", "支付日", "金額", "月份"]
        KIND_TO_COL = {
            "設計費": "設計費", "外管補助": "外管補助",
            "公共藝術": "公共藝術", "其他費用": "其他費用",
            "施工費": "施工費", "專管監造": "專管監造",
            "物調款": "物調款",
        }

        wide_rows = []

        df_custom_src = df_pay[df_pay["性質"].isin(
            ["設計費", "外管補助", "公共藝術", "其他費用"])].copy()
        for _, r in df_custom_src.iterrows():
            row = {c: "" for c in WIDE_COLS}
            row["期別"]   = r["期別"]
            row["支付日"] = r["支付日"]
            row["月份"]   = r["支付日"].strftime('%Y-%m')
            row["金額"]   = int(r["金額"])
            col_name = KIND_TO_COL.get(r["性質"])
            if col_name:
                row[col_name] = int(r["金額"])
            wide_rows.append(row)

        df_scurve_src = df_pay[df_pay["性質"].isin(
            ["施工費", "專管監造", "物調款"])].copy()
        if not df_scurve_src.empty:
            df_scurve_src["月份標籤"] = df_scurve_src["期別"].str.extract(r'工程估驗\s+(\d{4}/\d{2})')[0]
            for (month_label, pay_date), grp in df_scurve_src.groupby(["月份標籤", "支付日"]):
                row = {c: "" for c in WIDE_COLS}
                row["期別"]   = f"工程估驗 {month_label}"
                row["支付日"] = pay_date
                row["月份"]   = pay_date.strftime('%Y-%m')
                total_amt = 0
                for _, r in grp.iterrows():
                    col_name = KIND_TO_COL.get(r["性質"])
                    if col_name:
                        row[col_name] = int(r["金額"])
                        total_amt += int(r["金額"])
                row["金額"] = total_amt
                wide_rows.append(row)

        df_wide = pd.DataFrame(wide_rows, columns=WIDE_COLS)
        if not df_wide.empty:
            df_wide = df_wide.sort_values("支付日").reset_index(drop=True)

        tab1, tab2, tab3, tab4 = st.tabs([
            "📊 每月支出趨勢",
            "📜 詳細金流明細",
            "📑 寬表總覽",
            "📅 工期情境總結",
        ])

        with tab1:
            fig_bar = go.Figure(data=[go.Bar(
                x=df_monthly['月份'], y=df_monthly['金額'],
                marker_color='#2ecc71',
                text=[f"{v/10000:,.0f}萬" for v in df_monthly['金額']],
                textposition='outside'
            )])
            fig_bar.update_layout(title="<b>各月預計付款金額 (元) - 含全部費用</b>",
                                  template="plotly_white", height=450)
            st.plotly_chart(fig_bar, use_container_width=True)

        with tab2:
            st.markdown("#### 🏗️ 工程款按月明細（S-curve 撥款）")
            st.caption(f"基數 = 施工費 + 專管監造 + 物調款 = **{scurve_base_total:,.0f}** 元　"
                       f"每月按 S-curve 進度比例同步撥付此三項（準備金不參與 S-curve 撥款）")

            if not df_scurve_pivot.empty:
                df_scurve_pivot["合約年度"] = df_scurve_pivot["月份"].apply(
                    lambda m: get_contract_year(pd.to_datetime(m + "-01"), contract_d)
                )
                for year in df_scurve_pivot["合約年度"].unique():
                    with st.expander(f"📅 {year}（工程款）", expanded=True):
                        df_y = df_scurve_pivot[df_scurve_pivot["合約年度"] == year].copy()
                        df_show = df_y.drop(columns=["合約年度"]).copy()
                        df_show["當月進度"] = df_show["當月進度"].apply(lambda x: f"+{x:.2f}%")
                        for col in ["施工費", "專管監造", "物調款", "當月合計"]:
                            df_show[col] = df_show[col].apply(lambda x: f"{int(x):,}")
                        st.table(df_show)
                        st.markdown(f"**💰 {year} 工程款小計： `{int(df_y['當月合計'].sum()):,}` 元**")
            else:
                st.info("無 S-curve 工程款資料")

            st.markdown("---")
            st.markdown("#### 📋 自訂期別費用（設計費／外管／公共藝術／其他）")
            df_custom = df_pay[df_pay["性質"].isin(["設計費", "外管補助", "公共藝術", "其他費用"])].copy()
            if not df_custom.empty:
                df_custom = df_custom.sort_values("支付日")
                df_custom["合約年度"] = df_custom["支付日"].apply(lambda x: get_contract_year(x, contract_d))
                for year in df_custom["合約年度"].unique():
                    with st.expander(f"📅 {year}（自訂費用）", expanded=True):
                        df_y = df_custom[df_custom["合約年度"] == year].copy()
                        df_y["金額(元)"] = df_y["金額"].apply(lambda x: f"{x:,}")
                        df_y["支付日"] = pd.to_datetime(df_y["支付日"]).dt.strftime('%Y-%m-%d')
                        st.table(df_y[["支付日", "期別", "性質", "金額(元)"]])
                        st.markdown(f"**💰 {year} 自訂費用小計： `{df_y['金額'].sum():,}` 元**")
            else:
                st.info("無自訂期別費用")

        with tab3:
            st.markdown("#### 📑 寬表總覽")
            st.caption("每期一列，金額橫向分布到對應費用欄位；工程估驗為合併列，三個欄位（施工費/專管監造/物調款）同時有值")

            if df_wide.empty:
                st.info("尚無任何期別資料")
            else:
                df_wide_disp = df_wide.copy()
                df_wide_disp["合約年度"] = df_wide_disp["支付日"].apply(
                    lambda x: get_contract_year(x, contract_d))

                amount_cols = ["設計費", "外管補助", "公共藝術", "其他費用",
                               "專管監造", "物調款", "施工費", "金額"]

                def _fmt(v):
                    if v == "" or v is None: return ""
                    try:
                        return f"{int(v):,}"
                    except Exception:
                        return str(v)

                for year in df_wide_disp["合約年度"].unique():
                    df_y = df_wide_disp[df_wide_disp["合約年度"] == year].copy()
                    with st.expander(f"📅 {year}", expanded=True):
                        df_show = df_y.drop(columns=["合約年度"]).copy()
                        df_show["支付日"] = pd.to_datetime(df_show["支付日"]).dt.strftime('%Y-%m-%d')
                        for c in amount_cols:
                            df_show[c] = df_show[c].apply(_fmt)
                        st.dataframe(df_show, use_container_width=True, hide_index=True)

                        yearly_sum = int(pd.to_numeric(df_y["金額"], errors='coerce').fillna(0).sum())
                        st.markdown(f"**💰 {year} 合計： `{yearly_sum:,}` 元**")

                total_sum = int(pd.to_numeric(df_wide["金額"], errors='coerce').fillna(0).sum())
                st.markdown("---")
                st.markdown(f"### 🏁 全案總計： `{total_sum:,}` 元")

        with tab4:
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

                # ===== 🆕 將 S-Curve 圖轉成 PNG (bytes) =====
                # 說明：先把 Plotly 圖匯出為 PNG，若 kaleido 未安裝則給予提示
                scurve_png_bytes = None
                scurve_png_error = None
                try:
                    # 建立一個「匯出用」的 figure：移除 hovertemplate，放大尺寸，標題改單行
                    fig_export = go.Figure(fig)  # 複製
                    fig_export.update_layout(
                        width=1400,
                        height=700,
                        title=f"<b>{target_case_name} S-Curve 進度預測與風險分析</b>",
                        font=dict(family="Microsoft JhengHei, Arial, sans-serif", size=13),
                    )
                    scurve_png_bytes = fig_export.to_image(format="png", scale=2, engine="kaleido")
                except Exception as e:
                    scurve_png_error = str(e)

                # ===== 匯出多工作表 Excel =====
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    money_fmt = workbook.add_format({'num_format': '#,##0'})
                    pct_fmt   = workbook.add_format({'num_format': '0.00'})
                    bold_fmt  = workbook.add_format({'bold': True, 'num_format': '#,##0', 'bg_color': '#FFF3CD'})
                    title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'left'})
                    note_fmt  = workbook.add_format({'italic': True, 'font_color': '#666666'})

                    # ===== 🆕 Sheet 0: S-Curve 進度圖 =====
                    ws_chart = workbook.add_worksheet('S-Curve 進度圖')
                    writer.sheets['S-Curve 進度圖'] = ws_chart
                    ws_chart.set_column('A:A', 2)   # 左側留白
                    ws_chart.write('B2', f'{target_case_name} S-Curve 進度預測與風險分析', title_fmt)
                    ws_chart.write('B3', f'預測基準工期：{manual_dur} 天　｜　修正倍率：{env_ratio:.2f}　｜　模擬次數：{num_sims}', note_fmt)

                    if scurve_png_bytes is not None:
                        ws_chart.insert_image(
                            'B5', 'scurve.png',
                            {
                                'image_data': io.BytesIO(scurve_png_bytes),
                                'x_scale'   : 0.55,   # 縮放比例可調
                                'y_scale'   : 0.55,
                                'object_position': 1,
                            }
                        )
                    else:
                        # kaleido 未安裝或轉圖失敗 → 寫提示訊息取代圖
                        warn_fmt = workbook.add_format({'font_color': '#C00000', 'bold': True})
                        ws_chart.write('B5', '⚠️ 無法生成 S-Curve 圖片', warn_fmt)
                        ws_chart.write('B6', '請在終端機執行：pip install -U kaleido')
                        if scurve_png_error:
                            ws_chart.write('B7', f'錯誤訊息：{scurve_png_error}')

                    # ===== Sheet 1: 寬表總覽 =====
                    export_wide = df_wide.copy()
                    if not export_wide.empty:
                        export_wide["支付日"] = pd.to_datetime(export_wide["支付日"]).dt.strftime('%Y-%m-%d')
                        amt_cols = ["設計費", "外管補助", "公共藝術", "其他費用",
                                    "專管監造", "物調款", "施工費", "金額"]
                        for c in amt_cols:
                            export_wide[c] = pd.to_numeric(export_wide[c], errors='coerce').fillna(0).astype(int)
                        total_row = {c: "" for c in WIDE_COLS}
                        total_row["期別"] = "總計"
                        for c in amt_cols:
                            total_row[c] = int(export_wide[c].sum())
                        export_wide = pd.concat([export_wide, pd.DataFrame([total_row])], ignore_index=True)
                    export_wide.to_excel(writer, sheet_name='寬表總覽', index=False)

                    # ===== Sheet 2: 按月工程款 =====
                    export_scurve = df_scurve_pivot.drop(columns=["合約年度"], errors='ignore').copy()
                    if not export_scurve.empty:
                        export_scurve["當月進度"] = export_scurve["當月進度"].apply(lambda x: f"+{x:.2f}%")
                        total_row = {"月份": "總計", "當月進度": ""}
                        for c in ["施工費", "專管監造", "物調款", "當月合計"]:
                            total_row[c] = int(export_scurve[c].sum())
                        export_scurve = pd.concat([export_scurve, pd.DataFrame([total_row])], ignore_index=True)
                    export_scurve.to_excel(writer, sheet_name='按月工程款', index=False)

                    # ===== Sheet 3: 費用拆分 =====
                    fee_export_df = pd.DataFrame([
                        {"項目": "決標金額",             "比例(%)": "",              "金額(元)": total_p,         "基數": "",           "備註": "本案決標總金額"},
                        {"項目": "設計費",               "比例(%)": design_pct,      "金額(元)": design_f,        "基數": "決標金額",   "備註": "可調（自訂期別撥款）"},
                        {"項目": "統包施工費",           "比例(%)": "",              "金額(元)": const_p,         "基數": "",           "備註": "決標金額 − 設計費"},
                        {"項目": "專管及監造服務費",     "比例(%)": pcm_pct,         "金額(元)": pcm_fee,         "基數": "統包施工費", "備註": "含耐震特別監督，可調（S-curve 撥款）"},
                        {"項目": "準備金",               "比例(%)": 2.00,            "金額(元)": reserve_fee,     "基數": "統包施工費", "備註": "鎖定（不參與 S-curve 撥款）"},
                        {"項目": "物調款",               "比例(%)": price_adj_pct,   "金額(元)": price_adj_fee,   "基數": "統包施工費", "備註": "可調（S-curve 撥款）"},
                        {"項目": "外管補助費",           "比例(%)": 1.00,            "金額(元)": external_fee,    "基數": "統包施工費", "備註": "鎖定（自訂期別撥款）"},
                        {"項目": "公共藝術",             "比例(%)": 1.00,            "金額(元)": public_art_fee,  "基數": "統包施工費", "備註": "鎖定（自訂期別撥款）"},
                        {"項目": "其他費用",             "比例(%)": "",              "金額(元)": other_fee,       "基數": "",           "備註": "手動輸入（自訂期別撥款）"},
                        {"項目": "─────────",           "比例(%)": "",              "金額(元)": "",              "基數": "",           "備註": ""},
                        {"項目": "S-curve 基數",         "比例(%)": "",              "金額(元)": scurve_base_total, "基數": "",         "備註": "施工 + 專管 + 物調（不含準備金）"},
                        {"項目": "各項費用合計",         "比例(%)": "",              "金額(元)": total_all_fees,  "基數": "",           "備註": "設計費+專管+準備金+物調+外管+公共藝術+其他"},
                    ])
                    fee_export_df.to_excel(writer, sheet_name='費用拆分', index=False)

                    # ===== 格式設定 =====
                    ws_fee = writer.sheets['費用拆分']
                    ws_fee.set_column('A:A', 22)
                    ws_fee.set_column('B:B', 10, pct_fmt)
                    ws_fee.set_column('C:C', 18, money_fmt)
                    ws_fee.set_column('D:D', 14)
                    ws_fee.set_column('E:E', 30)

                    ws_sc = writer.sheets['按月工程款']
                    ws_sc.set_column('A:A', 12)
                    ws_sc.set_column('B:B', 12)
                    ws_sc.set_column('C:F', 16, money_fmt)
                    if not export_scurve.empty:
                        last_row = len(export_scurve)
                        ws_sc.set_row(last_row, None, bold_fmt)

                    ws_wide = writer.sheets['寬表總覽']
                    ws_wide.set_column('A:A', 24)
                    ws_wide.set_column('B:H', 14, money_fmt)
                    ws_wide.set_column('I:I', 13)
                    ws_wide.set_column('J:J', 16, money_fmt)
                    ws_wide.set_column('K:K', 10)
                    if not export_wide.empty:
                        last_row_w = len(export_wide)
                        ws_wide.set_row(last_row_w, None, bold_fmt)

                    # ===== 🆕 將 "S-Curve 進度圖" 排到第一個分頁 =====
                    workbook.worksheets_objs.sort(
                        key=lambda s: 0 if s.name == 'S-Curve 進度圖' else 1
                    )

                # 若轉圖失敗，畫面上提醒使用者
                if scurve_png_bytes is None:
                    st.warning(f"⚠️ S-Curve 圖片產生失敗，Excel 仍可下載但圖片頁為空白。\n"
                               f"請執行 `pip install -U kaleido`。\n\n錯誤：{scurve_png_error}")

                st.download_button("📥 下載預測報表 (.xlsx)", data=buffer.getvalue(),
                                   file_name=f"{target_case_name}_預測.xlsx")

    else:
        st.error(f"找不到工作表「{target_case_name}」")
else:
    st.info("💡 請上傳檔案以開始分析。系統將自動從歷史樣本提取廠商績效。")
