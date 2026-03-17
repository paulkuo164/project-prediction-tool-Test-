# --- 📈 3. 圖表渲染 (驗證版：含百分比顯示) ---
        def to_dates(curve): return [start_dt + timedelta(days=int(d * env_ratio)) for d in curve]
        
        u_days = (np.concatenate([target_df["天數"].values, mean_c])) * env_ratio
        u_prog = np.concatenate([target_df["累計_norm"].values, prog_steps])
        s_idx = np.argsort(u_days); u_days, u_prog = u_days[s_idx], u_prog[s_idx]

        # 準備 Hover 數據
        hover_custom_data = []
        for i, d in enumerate(mean_c):
            # 1. 取得當前點的平均進度 (y軸值)
            mean_prog = prog_steps[i]
            
            # 2. 透過插值找出在「同一天 d」時，樂觀與悲觀路徑對應的進度
            # 注意：p10, p90 陣列儲存的是「達到某進度所需的天數」
            p10_prog = np.interp(d, p10, prog_steps)
            p90_prog = np.interp(d, p90, prog_steps)
            
            # 3. 計算預估支付金額 (以 Mean 為基準)
            prev_d = max(0, d - 30)
            prog_now = np.interp(d * env_ratio, u_days, u_prog)
            prog_prev = np.interp(prev_d * env_ratio, u_days, u_prog)
            pay_amt = int(const_p * (prog_now - prog_prev) / 100)
            
            # 4. 計算價差金額 (P10進度 - P90進度)
            risk_gap_amt = int(const_p * (p10_prog - p90_prog) / 100)
            
            # 儲存所有要顯示在鼠標上的變數
            hover_custom_data.append([
                f"{pay_amt:,} 元",       # [0] 預估支付
                f"{risk_gap_amt:,} 元",   # [1] 價差金額
                f"{p10_prog:.2f}%",      # [2] 樂觀進度
                f"{mean_prog:.2f}%",     # [3] 平均進度
                f"{p90_prog:.2f}%"       # [4] 悲觀進度
            ])

        fig = go.Figure()
        
        # 區間層
        fig.add_trace(go.Scatter(x=to_dates(p10)+to_dates(p90)[::-1], y=prog_steps.tolist()+prog_steps[::-1].tolist(), fill='toself', fillcolor='rgba(149,165,166,0.1)', line=dict(color='rgba(255,255,255,0)'), name='90% 信賴區間', hoverinfo='skip'))
        fig.add_trace(go.Scatter(x=to_dates(p15)+to_dates(p85)[::-1], y=prog_steps.tolist()+prog_steps[::-1].tolist(), fill='toself', fillcolor='rgba(241,196,15,0.15)', line=dict(color='rgba(255,255,255,0)'), name='70% 信賴區間', hoverinfo='skip'))
        fig.add_trace(go.Scatter(x=to_dates(p25)+to_dates(p75)[::-1], y=prog_steps.tolist()+prog_steps[::-1].tolist(), fill='toself', fillcolor='rgba(46,134,193,0.2)', line=dict(color='rgba(255,255,255,0)'), name='50% 信賴區間', hoverinfo='skip'))
        
        # 主預測線
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
        
        fig.update_layout(
            title=f"<b>{target_case_name} S-Curve 進度預測與風險分析</b>",
            hovermode="x unified",
            template="plotly_white",
            height=650,
            legend=dict(orientation="h", y=1.1)
        )
        st.plotly_chart(fig, use_container_width=True)
