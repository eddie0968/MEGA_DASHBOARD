# -*- coding: utf-8 -*-
# 兆豐金控｜業務視覺化與同業差異分析（可直接運作）
# 使用方式：
# 1) 安裝依賴：pip install streamlit pandas plotly openpyxl numpy
# 2) 執行：streamlit run streamlit_megabank_dashboard.py
# 3) 將 Excel 放在同資料夾（檔名：14Jia-Jin-Kong-Zhu-Yao-Bi-Lu-final-3.xlsx），或於側欄上傳

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

from io import BytesIO

# =========================
# 資料讀取與清理
# =========================
def load_data(xlsx_bytes_or_path, sheet_name='14家金控主要比率'):
    """讀取附件Excel並萃取所有指標（自動轉數字、標示公股/民營）。"""
    if isinstance(xlsx_bytes_or_path, (bytes, bytearray)):
        df = pd.read_excel(BytesIO(xlsx_bytes_or_path), sheet_name=sheet_name)
    else:
        df = pd.read_excel(xlsx_bytes_or_path, sheet_name=sheet_name)

    # 依附件固定結構擷取第1~26列為資料區
    data = df.iloc[1:26].copy()
    
    # Define all possible columns that the app can handle
    all_possible_columns = [
        '公司', 'ROE_202506', 'ROE_2024', '消金貸款市占率', '信用卡流通卡市占率', '信用卡消費市占率',
        '信託商品市占率', '聯貸案市占率', '公民營放款市占率', '中小企放款市占率', '保證餘額市占率', '衍生性商品餘額市占率',
        '進出口信用狀市占率', '存款餘額市占率', '外匯存款餘額市占率', '臺股經紀市占率','複委託市占率', '融資市占率', 
        '股權承銷IPO送件數市占率', '股權承銷SPO送件數', 'CP2保證市占率', 'CP2承銷市占率', '票券買賣市占率', '債券買賣市占率',
        '保費收入市占率', '公募基金市占率'
    ]
    
    # Gracefully handle files with fewer columns than expected by using only the available columns
    actual_cols = data.shape[1]
    cols_to_assign = all_possible_columns[:actual_cols]
    data.columns = cols_to_assign

    # 數值欄轉數字
    for col in data.columns[1:]:
        data[col] = pd.to_numeric(data[col], errors='coerce')

    # 資料清理：移除公司名稱中的特定文字
    data['公司'] = data['公司'].str.replace('(註1)', '', regex=False).str.strip()
    # 移除無效資料行
    data = data[data['公司'] != '單位%'].copy()

    # 公股/民營預設清單（可於側欄覆寫）
    default_public = ['華南金','第一金','合庫金','兆豐金','臺灣金控']
    data['類別'] = data['公司'].apply(lambda x: '公股' if x in default_public else '民營')

    return data

st.set_page_config(page_title="兆豐金控業務視覺化儀表板", page_icon="📊", layout="wide")
SUBSIDIARY_GROUPS = {
    '銀行': [
        'ROE_202506', 'ROE_2024', '消金貸款市占率', '信用卡流通卡市占率', '信用卡消費市占率',
        '信託商品市占率', '聯貸案市占率', '公民營放款市占率', '中小企放款市占率', '保證餘額市占率',
        '衍生性商品餘額市占率', '進出口信用狀市占率', '存款餘額市占率', '外匯存款餘額市占率'
    ],
}

# Load the main data
data_df = load_data("14家金控主要比率final 3.xlsx")


# =========================
# 資料讀取與清理
# =========================
def load_data(xlsx_bytes_or_path, sheet_name='14家金控主要比率'):
    """讀取附件Excel並萃取所有指標（自動轉數字、標示公股/民營）。"""
    if isinstance(xlsx_bytes_or_path, (bytes, bytearray)):
        df = pd.read_excel(BytesIO(xlsx_bytes_or_path), sheet_name=sheet_name)
    else:
        df = pd.read_excel(xlsx_bytes_or_path, sheet_name=sheet_name)

    # 依附件固定結構擷取第1~26列為資料區
    data = df.iloc[1:26].copy()
    
    # Define all possible columns that the app can handle
    all_possible_columns = [
        '公司', 'ROE_202506', 'ROE_2024', '消金貸款市占率', '信用卡流通卡市占率', '信用卡消費市占率',
        '信託商品市占率', '聯貸案市占率', '公民營放款市占率', '中小企放款市占率', '保證餘額市占率', '衍生性商品餘額市占率',
        '進出口信用狀市占率', '存款餘額市占率', '外匯存款餘額市占率', '臺股經紀市占率','複委託市占率', '融資市占率', 
        '股權承銷IPO送件數市占率', '股權承銷SPO送件數', 'CP2保證市占率', 'CP2承銷市占率', '票券買賣市占率', '債券買賣市占率',
        '保費收入市占率', '公募基金市占率'
    ]
    
    # Gracefully handle files with fewer columns than expected by using only the available columns
    actual_cols = data.shape[1]
    cols_to_assign = all_possible_columns[:actual_cols]
    data.columns = cols_to_assign

    # 數值欄轉數字
    for col in data.columns[1:]:
        data[col] = pd.to_numeric(data[col], errors='coerce')

    # 資料清理：移除公司名稱中的特定文字
    data['公司'] = data['公司'].str.replace('(註1)', '', regex=False).str.strip()
    # 移除無效資料行
    data = data[data['公司'] != '單位%'].copy()

    # 公股/民營預設清單（可於側欄覆寫）
    default_public = ['華南金','第一金','合庫金','兆豐金','臺灣金控']
    data['類別'] = data['公司'].apply(lambda x: '公股' if x in default_public else '民營')

    return data

# 排名計算（降序：值越大名次越前）
def add_ranking(df, indicators):
    out = df.copy()
    for ind in indicators:
        rank_col = f"{ind}_排名"
        out[rank_col] = out[ind].rank(ascending=False, method='min')
    return out

# 百分位與差異輔助
def add_percentile(df, indicators):
    out = df.copy()
    for ind in indicators:
        pcol = f"{ind}_百分位"
        ser = out[ind]
        out[pcol] = ser.rank(pct=True)
    return out

# =========================
# 視覺化元件
# =========================
def overview_charts(bank, indicators, sort_by, category_filter, peer_metric, sub_type):
    data = bank.copy()
    if category_filter in ['公股','民營']:
        data = data[data['類別'] == category_filter]

    # 為了凸顯兆豐，增加一個欄位
    data['plot_group'] = data.apply(
        lambda row: '兆豐金' if row['公司'] == '兆豐金' else row['類別'],
        axis=1
    )

    # 1) 分組條形圖（按所選指標，以公司為x、數值為y、顏色分組為公股/民營）
    df_long = data.melt(id_vars=['公司', 'plot_group'], value_vars=indicators, var_name='指標', value_name='數值')
    if sort_by in indicators:
        order = data.sort_values(sort_by, ascending=False)['公司']
        category_orders = {'公司': list(order)}
    else:
        category_orders = None
    fig = px.bar(
        df_long, x='公司', y='數值', color='plot_group', barmode='group', facet_row='指標',
        category_orders=category_orders,
        color_discrete_map={'公股':'#127C90','民營':'#D64550', '兆豐金': '#F4B000'}
    )
    
    # Add average/median line
    if indicators:
        peer_values = bank[bank['公司'] != '兆豐金'][indicators]
        agg_values = peer_values.mean() if peer_metric == '平均數' else peer_values.median()
        
        # Facet rows are ordered from top to bottom based on `indicators` list
        for i, indicator in enumerate(indicators):
            line_color = "red" if peer_metric == '平均數' else "orange"
            fig.add_hline(
                y=agg_values[indicator],
                line_dash="dot",
                line_color=line_color,
                row=len(indicators) - i, col=1
            )

    fig.update_layout(
        title_text=f"{sub_type}業務指標比較圖<br><sup>單位：%</sup><br><sup>同業比較基準: {peer_metric}</sup>",
        height=650,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.4, 
            xanchor="right",
            x=1
        ),
        margin=dict(r=150) # 增加右側邊距以避免文字截斷
    )
    
    # 將指標標籤改為水平並放置在右側
    def update_facet_annotations(a):
        if a.text.startswith('指標='):
            a.text = a.text.split('=')[1]
            a.x = 1.01  # Adjust x-position to the right
            a.xanchor = 'left' # Anchor to the left of the x-position
            a.textangle = 0
    fig.for_each_annotation(update_facet_annotations)

    # 顯示 y 軸的刻度標籤，並清除y軸標題
    fig.update_yaxes(showticklabels=True, title_text="")
    
    return fig

def peer_comparison(bank, target, peers, indicators, category_filter):
    data = bank.copy()
    if category_filter in ['公股','民營']:
        data = data[data['類別'] == category_filter]

    df_sel = data[data['公司'].isin([target] + peers)].copy()
    if df_sel.empty:
        st.warning("未找到選定的公司/同業資料，請重新選擇。")
        return None, None

    # 分組條形圖（指標為facet）
    df_long = df_sel.melt(id_vars=['公司','類別'], value_vars=indicators, var_name='指標', value_name='數值')
    fig = px.bar(
        df_long, x='公司', y='數值', color='公司', barmode='group', facet_row='指標',
        color_discrete_sequence=px.colors.qualitative.Set2
    )
    fig.update_layout(title=f"{target} 與自選同業差異比較（{category_filter if category_filter in ['公股','民營'] else '全部'}）")
    st.plotly_chart(fig, use_container_width=True)

    # 差異表（以 target 為基準）
    base = bank[bank['公司'] == target][['公司'] + indicators].set_index('公司').iloc[0]
    diffs = []
    for p in peers:
        if p in df_sel['公司'].values:
            row = df_sel[df_sel['公司'] == p][['公司'] + indicators].set_index('公司').iloc[0]
            delta = (base - row).to_dict()
            delta['同業'] = p
            diffs.append(delta)
    diff_df = pd.DataFrame(diffs).set_index('同業') if diffs else pd.DataFrame()

    if not diff_df.empty:
        st.subheader("與同業差異（正值=目標較高，負值=目標較低）")
        st.dataframe(diff_df.style.background_gradient(cmap='RdYlGn'))
        csv = diff_df.to_csv().encode('utf-8-sig')
        st.download_button("下載差異表CSV", data=csv, file_name="peer_diff.csv", mime="text/csv")

    return fig, diff_df

def group_radar(bank, indicators, peer_metric):
    if peer_metric == '平均數':
        public_agg = bank[bank['類別']=='公股'][indicators].mean()
        private_agg = bank[bank['類別']=='民營'][indicators].mean()
    else:
        public_agg = bank[bank['類別']=='公股'][indicators].median()
        private_agg = bank[bank['類別']=='民營'][indicators].median()

    mega_series = bank[bank['公司']=='兆豐金'][indicators].iloc[0]

    plot_df = pd.DataFrame({
        '指標': indicators,
        '兆豐金': mega_series.values,
        f'公股 {peer_metric}': public_agg.values,
        f'民營 {peer_metric}': private_agg.values
    })

    fig = go.Figure()
    for col, color in [('兆豐金', '#0D4C92'), (f'公股 {peer_metric}', '#2BA84A'), (f'民營 {peer_metric}', '#E4572E')]:
        fig.add_trace(go.Scatterpolar(r=plot_df[col], theta=plot_df['指標'], fill='toself', name=col, opacity=0.6, line=dict(color=color)))

    fig.update_layout(title=f'公股/民營分組 vs 兆豐金 雷達圖 ({peer_metric})', polar=dict(radialaxis=dict(visible=True)))
    st.plotly_chart(fig, use_container_width=True)

    # 提供下載HTML
    html_bytes = fig.to_html(include_plotlyjs='cdn').encode('utf-8')
    st.download_button("下載雷達圖(HTML)", data=html_bytes, file_name="group_radar.html", mime="text/html")

    return fig

# =========================
# 介面
# =========================
st.title("📊 兆豐金控 業務視覺化與同業差異分析")
st.caption("金控競爭地位總覽、子公司同業差異與公民營分組地位比較")

# --- Sidebar UI and Data Loading ---
with st.sidebar:
    st.header("資料來源")
    use_builtin = st.toggle("資料來源:金控業比較分析表", value=True)
    uploaded = None
    if not use_builtin:
        uploaded = st.file_uploader("上傳最新Excel（14家金控主要比率）", type=["xlsx"])

    st.markdown("---")
    st.header("篩選/分組")
    category_filter = st.radio("公股/民營篩選", options=["全部","公股","民營"], index=0, horizontal=True)
    peer_metric = st.radio("同業比較基準", ["平均數", "中位數"], index=0, horizontal=True)

    st.markdown("---")
    st.header("資料設定")

    existing_cols = set(data_df.columns)
    filtered_subsidiary_groups = {
        group: [col for col in metrics if col in existing_cols]
        for group, metrics in SUBSIDIARY_GROUPS.items()
    }
    filtered_subsidiary_groups = {k: v for k, v in filtered_subsidiary_groups.items() if v}

    if not filtered_subsidiary_groups:
        st.error("檔案中未找到可識別的指標。")
        st.stop()

    sub_type = st.selectbox("選擇子公司類型", options=list(filtered_subsidiary_groups.keys()), index=0)

    st.markdown("---")
    st.header("指標設定")
    all_inds = filtered_subsidiary_groups[sub_type]
    default_inds = all_inds[:5] if len(all_inds) > 5 else all_inds
    
    selected_inds = st.multiselect(f"選擇 {sub_type} 指標", options=all_inds, default=["ROE_202506", "存款餘額市占率", "外匯存款餘額市占率"])
    sort_by = st.selectbox("公司排序依據", options=['(不排序)'] + selected_inds, index=0) # Default to '(不排序)'

st.markdown("---")

# 分頁
tab1, tab2, tab3, tab4 = st.tabs(["同業比較全覽", "同業金控比較分析", "公股/民營分組", "業務指標熱力圖"])

with tab1:
    st.subheader("同業比較全覽")
    if not selected_inds:
        st.warning("請在左側「指標設定」中至少選擇一個指標。 সন")
    else:
        # --- Section 1: Main bar chart (full width) ---
        fig_overview = overview_charts(data_df, selected_inds, sort_by if sort_by != '(不排序)' else selected_inds[0], category_filter, peer_metric, sub_type)
        st.plotly_chart(fig_overview, use_container_width=True)
        html_bytes_overview = fig_overview.to_html(include_plotlyjs='cdn').encode('utf-8')
        st.download_button("下載概覽圖(HTML)", data=html_bytes_overview, file_name="overview.html", mime="text/html", key="dl_overview")

        st.markdown("---")

        # --- Section 2: Ranking dataframe for comparison ---
        st.subheader("同業業務指標比較表")

        rank_df = add_ranking(data_df, selected_inds)
        show_cols = ['公司','類別'] + selected_inds
        
        def highlight_mega(row):
            if row['公司'] == '兆豐金':
                return ['background-color: #F4B000'] * len(row)
            else:
                return [''] * len(row)

        display_df = rank_df[show_cols]
        
        peer_values = data_df[data_df['公司'] != '兆豐金'][selected_inds]
        agg_series = peer_values.mean() if peer_metric == '平均數' else peer_values.median()
        agg_series.name = f"同業{peer_metric}"
        agg_df = pd.DataFrame(agg_series).T
        agg_df.reset_index(inplace=True)
        agg_df = agg_df.rename(columns={'index': '公司'})
        display_df = pd.concat([display_df, agg_df], ignore_index=True)

        format_dict = {}
        for col in selected_inds:
            format_dict[col] = "{:.2f}"

        def highlight_peer_metric(row):
            if str(row['公司']).startswith('同業'):
                return ['background-color: #800020; color: white'] * len(row)
            return [''] * len(row)

        display_df.index = np.arange(1, len(display_df) + 1)
        st.dataframe(display_df.style.apply(highlight_mega, axis=1).apply(highlight_peer_metric, axis=1).format(format_dict), use_container_width=True)
        
        st.markdown("---")

        # --- Section 3: Heatmaps side-by-side ---
        col_heat1, col_heat2 = st.columns(2)

        with col_heat1:
            # Business Metrics Heatmap (from original overview_charts)
            st.write("#### 業務指標熱力圖")
            data = data_df.copy()
            if category_filter in ['公股','民營']:
                data = data[data['類別'] == category_filter]
            heat_df = data.set_index('公司')[selected_inds]
            heat_df_filtered = heat_df.dropna(how='all')
            
            z = heat_df_filtered.values
            fig_heat = go.Figure(data=go.Heatmap(
                z=z,
                x=selected_inds,
                y=heat_df_filtered.index,
                colorscale='Blues',
                colorbar=dict(title='數值')
            ))
            fig_heat.update_layout(title="業務指標熱力圖", height=600)
            st.plotly_chart(fig_heat, use_container_width=True)
            html_bytes_heat = fig_heat.to_html(include_plotlyjs='cdn').encode('utf-8')
            st.download_button("下載業務指標熱力圖(HTML)", data=html_bytes_heat, file_name="heatmap.html", mime="text/html", key="dl_heatmap")

        with col_heat2:
            # Ranking Heatmap (from original tab4)
            st.write("#### 名次熱力圖")
            sort_key_indicator = sort_by if sort_by in selected_inds else selected_inds[0]
            sort_col = f"{sort_key_indicator}_排名"
            # rank_df was already computed above
            sorted_rank_df = rank_df.sort_values(by=sort_col, ascending=True)

            heat = sorted_rank_df.set_index('公司')[[f"{c}_排名" for c in selected_inds]]
            heat = heat.replace(np.inf, np.nan)
            
            figr = go.Figure(data=go.Heatmap(
                z=heat.values,
                x=heat.columns,
                y=heat.index,
                colorscale='Tealgrn',
                reversescale=True,
                colorbar=dict(title='名次 (1最好)'),
                text=heat.values,
                texttemplate="%{text:.0f}",
                textfont={"size":12}
            ))
            figr.update_layout(title=f"名次熱力圖 (依 {sort_key_indicator} 排序)", height=600)
            st.plotly_chart(figr, use_container_width=True)
            html_bytes_figr = figr.to_html(include_plotlyjs='cdn').encode('utf-8')
            st.download_button("下載名次熱力圖(HTML)", data=html_bytes_figr, file_name="rank_heatmap.html", mime="text/html", key="dl_rank_heatmap")

with tab2:
    st.subheader("同業金控比較分析")
    target = '兆豐金' # 預設目標公司為兆豐金
    peer_candidates = [c for c in data_df['公司'].tolist() if c != '兆豐金']
    peers = st.multiselect("選擇1~4家同業比較", options=peer_candidates, default=['中信金','玉山金','華南金','第一金'][:min(4,len(peer_candidates))])
    fig_peer, diff_df = peer_comparison(data_df, target, peers, selected_inds, category_filter)
    if fig_peer is not None:
        html_bytes = fig_peer.to_html(include_plotlyjs='cdn').encode('utf-8')
        st.download_button("下載同業比較(HTML)", data=html_bytes, file_name="peer_comparison.html", mime="text/html")

with tab3:
    st.subheader("公股/民營分組地位")
    fig_radar = group_radar(data_df, selected_inds, peer_metric)

with tab4:
    st.subheader("業務指標熱力圖")

    # Allow user to select indicators for the heatmap
    heatmap_indicators = st.multiselect(
        "選擇熱力圖指標",
        options=all_inds,
        default=["存款餘額市占率", "外匯存款餘額市占率"] # Default to specific indicators
    )

    if heatmap_indicators:
        for indicator in heatmap_indicators:
            # Apply category filter
            filtered_data_df = data_df.copy()
            if category_filter in ['公股', '民營']:
                filtered_data_df = filtered_data_df[filtered_data_df['類別'] == category_filter]

            # Prepare data for treemap for the current indicator
            df_for_single_heatmap = filtered_data_df[['公司', indicator]].copy()
            
            # Rename the indicator column to '數值' for treemap
            df_for_single_heatmap.rename(columns={indicator: '數值'}, inplace=True)

            # Calculate Mega Bank's value and rank from the unfiltered data_df
            mega_value = data_df[data_df['公司'] == '兆豐金'][indicator].iloc[0]
            
            # Calculate rank for all companies in the unfiltered data_df
            temp_rank_df = data_df[['公司', indicator]].copy()
            temp_rank_df['rank_col'] = temp_rank_df[indicator].rank(ascending=False, method='min')
            mega_rank = temp_rank_df[temp_rank_df['公司'] == '兆豐金']['rank_col'].iloc[0]
            
            # Calculate rank for df_for_single_heatmap (filtered data)
            df_for_single_heatmap['rank_col'] = df_for_single_heatmap['數值'].rank(ascending=False, method='min')

            # Add rank text for display on squares
            df_for_single_heatmap['rank_display'] = '#' + df_for_single_heatmap['rank_col'].fillna(0).astype(int).astype(str)

            num_companies = len(df_for_single_heatmap)
            
            # Create the treemap for the single indicator
            fig_treemap = px.treemap(
                df_for_single_heatmap,
                path=['公司'], # Only company as path
                values='數值',
                color='數值', # Color by '數值'
                color_continuous_scale=['#A8D8AD', '#FFFFCC', '#F08080'], # Softer Green for low, Yellow for mid, Red for high
                title=f"{indicator} 熱力圖", # Removed color explanation from title
                custom_data=['rank_display'] # Pass rank_display as custom data
            )
            # Add annotation for color explanation
            fig_treemap.add_annotation(
                text="數值越大顏色越紅，數值越小顏色越綠", # Updated text
                xref="paper", yref="paper",
                x=1, y=-0.15, # Position at bottom right
                showarrow=False,
                font=dict(size=10, color="gray"),
                xanchor="right", yanchor="auto"
            )
            fig_treemap.update_traces(
                texttemplate="%{label}<br>%{value:.2f}%<br>%{customdata[0]}", # Add company name, value, and rank
                textfont=dict(
                    family="Microsoft JhengHei, sans-serif", # 微軟正黑體
                    size=18, # Larger font size
                    color="black" # Ensure good contrast
                )
            )

            # Highlight Mega Bank with a white, thick border
            for i, d in enumerate(fig_treemap.data[0].labels):
                if d == '兆豐金':
                    fig_treemap.data[0].marker.line.width = [6 if label == '兆豐金' else 0 for label in fig_treemap.data[0].labels] # Even thicker border
                    fig_treemap.data[0].marker.line.color = ['white' if label == '兆豐金' else 'rgba(0,0,0,0)' for label in fig_treemap.data[0].labels] # White color
                    break
            
            # Calculate peer metric value
            peer_values_for_indicator = data_df[data_df['公司'] != '兆豐金'][indicator]
            if peer_metric == '平均數':
                peer_metric_value = peer_values_for_indicator.mean()
            else: # 中位數
                peer_metric_value = peer_values_for_indicator.median()

            # Calculate difference
            difference = mega_value - peer_metric_value

            # Create the treemap for the single indicator
            fig_treemap = px.treemap(
                df_for_single_heatmap,
                path=['公司'], # Only company as path
                values='數值',
                color='數值', # Color by '數值'
                color_continuous_scale=['#A8D8AD', '#FFFFCC', '#F08080'], # Softer Green for low, Yellow for mid, Red for high
                title=f"{indicator} 熱力圖", # Removed color explanation from title
                custom_data=['rank_display'] # Pass rank_display as custom data
            )
            # Add annotation for color explanation
            fig_treemap.add_annotation(
                text="數值越大顏色越紅，數值越小顏色越綠", # Updated text
                xref="paper", yref="paper",
                x=1, y=-0.15, # Position at bottom right
                showarrow=False,
                font=dict(size=10, color="gray"),
                xanchor="right", yanchor="auto"
            )
            fig_treemap.update_traces(
                texttemplate="%{label}<br>%{value:.2f}%<br>%{customdata[0]}", # Add company name, value, and rank
                textfont=dict(
                    family="Microsoft JhengHei, sans-serif", # 微軟正黑體
                    size=18, # Larger font size
                    color="black" # Ensure good contrast
                )
            )

            # Highlight Mega Bank with a white, thick border
            for i, d in enumerate(fig_treemap.data[0].labels):
                if d == '兆豐金':
                    fig_treemap.data[0].marker.line.width = [6 if label == '兆豐金' else 0 for label in fig_treemap.data[0].labels] # Even thicker border
                    fig_treemap.data[0].marker.line.color = ['white' if label == '兆豐金' else 'rgba(0,0,0,0)' for label in fig_treemap.data[0].labels] # White color
                    break
            
            st.markdown(f"<p style='font-size:24px; color:#FFFFFF;'><b>兆豐金 {indicator}</b>: {mega_value:.2f}% (在 {num_companies} 家金控中排名第 {int(mega_rank)} 名) | 同業{peer_metric}: {peer_metric_value:.2f}% (差異: {difference:+.2f}%)</p>", unsafe_allow_html=True)

            st.plotly_chart(fig_treemap, use_container_width=True, key=f"treemap_{indicator}") # Added unique key

            

            # Download button for the treemap
            html_bytes_treemap = fig_treemap.to_html(include_plotlyjs='cdn').encode('utf-8')
            st.download_button(
                f"下載 {indicator} 熱力圖(HTML)",
                data=html_bytes_treemap,
                file_name=f"{indicator}_heatmap.html",
                mime="text/html",
                key=f"dl_treemap_{indicator}" # Unique key for each download button
            )
            st.markdown("---") # Separator between heatmaps
    else:
        st.info("請選擇至少一個業務指標以顯示熱力圖。")
