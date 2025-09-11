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
import base64
from pathlib import Path
from streamlit_sortables import sort_items 

# =========================
# Logo Utility Function
# =========================
@st.cache_data
def get_logo_data_uri(company_name):
    """Looks for a logo file, encodes it and returns a base64 data URI."""
    logo_extensions = ['.png', '.jpg', '.jpeg', '.jfif']
    # Assume logos are in a 'logos' subdirectory
    logo_dir = Path('logos')
    if not logo_dir.is_dir():
        return None
    
    for ext in logo_extensions:
        logo_file = logo_dir / f"{company_name}{ext}"
        if logo_file.is_file():
            try:
                encoded = base64.b64encode(logo_file.read_bytes()).decode()
                return f"data:image/{ext.replace('.', '')};base64,{encoded}"
            except Exception as e:
                # st.warning(f"Could not read or encode logo for {company_name}: {e}")
                return None
    return None

# =========================
# 資料讀取與清理
# =========================
@st.cache_data
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
    default_public = ['華南金控','第一金控','合庫金控','兆豐金控','臺灣金控']
    data['類別'] = data['公司'].apply(lambda x: '公股' if x in default_public else '民營')

    return data

@st.cache_data
def load_dynamic_data(xlsx_bytes_or_path, sheet_name='ROE'): # Sheet name is 'ROE'
    """讀取歷年走勢Excel並萃取資料。"""
    if isinstance(xlsx_bytes_or_path, (bytes, bytearray)):
        df = pd.read_excel(BytesIO(xlsx_bytes_or_path), sheet_name=sheet_name)
    else:
        df = pd.read_excel(xlsx_bytes_or_path, sheet_name=sheet_name)

    # Assuming first column is company name, and subsequent columns are ROE for different years
    # Convert all columns except the first to numeric
    for col in df.columns[1:]:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    
    # Fill NaN values with 0 for plotting purposes, or handle as appropriate
    df = df.fillna(0)

    return df

@st.cache_data
def load_dynamic2_data(file_path):
    """Loads all sheets from Dynamic2.xlsx and returns them as a dict of dataframes, plus sheet names."""
    try:
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        data_sheets = {}
        for sheet in sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            # Basic cleaning
            df.rename(columns={df.columns[0]: '公司'}, inplace=True)
            for col in df.columns[1:]:
                df[col] = pd.to_numeric(df[col], errors='coerce')
            df = df.fillna(0)
            data_sheets[sheet] = df
        return data_sheets, sheet_names
    except FileNotFoundError:
        return None, []

st.set_page_config(page_title="兆豐金控業務視覺化儀表板", page_icon="兆豐金.jfif", layout="wide")
SUBSIDIARY_GROUPS = {
    '銀行': [
        '消金貸款市占率', '信用卡流通卡市占率', '信用卡消費市占率',
        '信託商品市占率', '聯貸案市占率', '公民營放款市占率', '中小企放款市占率', '保證餘額市占率',
        '衍生性商品餘額市占率', '進出口信用狀市占率', '存款餘額市占率', '外匯存款餘額市占率'
    ],
    '證券': [
        '臺股經紀市占率', '複委託市占率', '融資市占率', '股權承銷IPO送件數市占率', '股權承銷SPO送件數'
    ],
    '票券': [
        'CP2保證市占率', 'CP2承銷市占率', '票券買賣市占率', '債券買賣市占率'
    ],
    '保險': [
        '保費收入市占率'
    ],
    '投信': [
        '公募基金市占率'
    ]
}

# Load the main data
data_df = load_data("14家金控主要比率final 3.xlsx")

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
    data['plot_group'] = np.where(data['公司'] == '兆豐金控', '兆豐金控', data['類別'])

    # 1) 分組條形圖（按所選指標，以公司為x、數值為y、顏色分組為公股/民營）
    df_long = data.melt(id_vars=['公司', 'plot_group'], value_vars=indicators, var_name='指標', value_name='數值')
    if sort_by in indicators:
        order = data.sort_values(sort_by, ascending=False)['公司']
        category_orders = {'公司': list(order)}
    else:
        category_orders = None
    fig = px.bar(
        df_long, x='數值', y='公司', color='plot_group', orientation='h', facet_row='指標',
        category_orders=category_orders,
        color_discrete_map={'公股':'#127C90','民營':'#D64550', '兆豐金控': '#F4B000'}
    )
    # Add average and median lines with legend entries
    if indicators:
        peer_data_for_agg = data[~data['公司'].isin(['兆豐金控'])][indicators] # Removed '其他公司' from exclusion
        
        mean_values = peer_data_for_agg.mean()
        median_values = peer_data_for_agg.median()
        
        for i, indicator in enumerate(indicators):
            # Add Mean Line
            fig.add_vline(
                x=mean_values[indicator],
                line_dash="dot",
                line_color="red",
                row=len(indicators) - i, col=1
            )
            # Add Median Line
            fig.add_vline(
                x=median_values[indicator],
                line_dash="dash",
                line_color="orange",
                row=len(indicators) - i, col=1
            )
        
        # Add dummy traces for legend entries
        fig.add_trace(go.Scatter(
            x=[None], y=[None],
            mode='lines',
            line=dict(color="red", dash="dot"),
            name='同業平均',
            showlegend=True
        ))
        fig.add_trace(go.Scatter(
            x=[None], y=[None],
            mode='lines',
            line=dict(color="orange", dash="dash"),
            name='同業中位數',
            showlegend=True
        ))

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

def create_radar_chart_for_tab1(bank, indicators, category_filter, peer_metric):
    if len(indicators) < 3:
        st.warning("雷達圖至少需要3個業務指標才能繪製。\n")
        return None

    data = bank.copy()
    
    # Get Mega Bank's data from the unfiltered bank DataFrame
    mega_series = bank[bank['公司'] == '兆豐金控'][indicators].iloc[0]

    if category_filter in ['公股','民營']:
        data = data[data['類別'] == category_filter]

    # Get peer data (excluding Mega Bank)
    peer_data = data[data['公司'] != '兆豐金控'][indicators]

    # Calculate peer aggregate (mean or median)
    if peer_metric == '平均數':
        peer_agg = peer_data.mean()
    else:
        peer_agg = peer_data.median()

    # Prepare data for plotting
    plot_df = pd.DataFrame({
        '指標': indicators,
        '兆豐金控': mega_series.values,
        f'同業 {peer_metric}': peer_agg.values
    })

    fig = go.Figure()
    # 兆豐金控
    fig.add_trace(go.Scatterpolar(
        r=plot_df['兆豐金控'], 
        theta=plot_df['指標'], 
        fill='toself', 
        name='兆豐金控', 
        opacity=0.9,
        line=dict(color='#F4B000', width=4),
        mode='lines+markers+text',
        text=plot_df['兆豐金控'].round(2),
        textfont=dict(color='white', size=10),
        marker=dict(color='white', size=6, symbol='diamond')
    ))
    # 同業
    fig.add_trace(go.Scatterpolar(
        r=plot_df[f'同業 {peer_metric}'], 
        theta=plot_df['指標'], 
        fill='toself', 
        name=f'同業 {category_filter if category_filter in ["公股","民營"] else ""} {peer_metric}', # Reflect category filter in legend 
        opacity=0.4, 
        line=dict(color='#2BA84A')
    ))

    fig.update_layout(
        title=f'兆豐金控 vs 同業 {peer_metric} 雷達圖 ({category_filter if category_filter in ["公股","民營"] else "全部"})',
        polar=dict(radialaxis=dict(visible=True))
    )
    return fig

def peer_comparison(bank, target, peers, indicators, category_filter):
    # Calculate base values for the target company from the unfiltered bank DataFrame
    target_data = bank[bank['公司'] == target]
    if target_data.empty:
        st.error(f"目標公司 '{target}' 不在資料中，請檢查資料或目標公司名稱。")
        return None, None
    base_values = target_data[indicators].iloc[0]

    data = bank.copy()
    if category_filter in ['公股','民營']:
        data = data[data['類別'] == category_filter]

    # Ensure target is always in df_sel for comparison, even if filtered out by category
    companies_to_include = [target] + peers
    df_sel = data[data['公司'].isin(companies_to_include)].copy()

    if df_sel.empty:
        st.warning("未找到選定的公司/同業資料，請重新選擇.\n")
        return None, None

    # Melt the DataFrame to long format
    df_long = df_sel.melt(id_vars=['公司', '類別'], value_vars=indicators, var_name='指標', value_name='數值')
    
    # Create a DataFrame for the base values for merging
    base_long = base_values.to_frame(name='基準數值').reset_index().rename(columns={'index': '指標'})

    # Merge to get base values
    df_merged = pd.merge(df_long, base_long, on='指標')

    # Calculate difference
    df_merged['差異數'] = df_merged['數值'] - df_merged['基準數值']

    # Calculate public and private averages of differences from ALL companies
    all_long = data.melt(id_vars=['公司', '類別'], value_vars=indicators, var_name='指標', value_name='數值')
    all_merged = pd.merge(all_long, base_long, on='指標')
    all_merged['差異數'] = all_merged['數值'] - all_merged['基準數值']

    public_diff_avg = all_merged[all_merged['類別'] == '公股'].groupby('指標')['差異數'].mean().reset_index()
    public_diff_avg['公司'] = '公股平均'
    public_diff_avg['類別'] = '公股'

    private_diff_avg = all_merged[all_merged['類別'] == '民營'].groupby('指標')['差異數'].mean().reset_index()
    private_diff_avg['公司'] = '民營平均'
    private_diff_avg['類別'] = '民營'

    # Combine all data for plotting
    plot_data = pd.concat([df_merged, public_diff_avg, private_diff_avg], ignore_index=True)

    # Calculate dynamic height and spacing based on number of indicators
    num_indicators = len(indicators)
    if num_indicators > 1:
        facet_spacing = min(0.4, 1 / (num_indicators - 1) * 0.8)
    else:
        facet_spacing = 0
    chart_height = max(400, num_indicators * (200 + 150 * facet_spacing)) # Adjust height based on spacing

    # Define custom colors for specific companies/groups
    color_map = {
        target: '#F4B000', # Mega Financial Holding in orange
        '公股平均': '#2BA84A', # Public average in green
        '民營平均': '#E4572E' # Private average in red
    }
    # Assign other companies a default color, e.g., a shade of blue
    for company in plot_data['公司'].unique():
        if company not in color_map:
            color_map[company] = '#127C90' # Default blue for other companies

    # Create the dot plot
    fig = px.scatter(
        plot_data,
        x='差異數',
        y='公司',
        color='公司', # Color by company
        color_discrete_map=color_map,
        facet_row='指標', # Change to facet_row for vertical stacking
        facet_row_spacing=facet_spacing, # Increase spacing between rows
        height=chart_height, # Use dynamic height
        title=f"{target} 與所選同業差異比較",
        labels={'差異數': '與兆豐金控的差異數'},
        hover_data={'差異數': ':.2f', '公司': True, '指標': True},
        text='公司' # Add company name as text
    )

    fig.update_traces(marker=dict(size=12, line=dict(width=1, color='DarkSlateGrey')))

    # Highlight target company with a larger marker and remove its text label
    fig.for_each_trace(lambda trace:
        trace.update(marker_size=15, marker_symbol='star', text=None) if trace.name == target else (trace.update(marker_symbol='circle') if trace.name in ['公股平均', '民營平均'] else None)
    )

    # Adjust text position for all traces (except target if text=None works)
    fig.update_traces(textposition='top center')

    # Add vertical line at 0 for reference
    fig.add_vline(x=0, line_width=1, line_dash="dash", line_color="gray")

    # Increase font sizes for better readability
    fig.update_layout(
        title_font_size=24,
        font=dict(size=14), # General font size for axes, ticks, legend
        xaxis=dict(title_font=dict(size=16), tickfont=dict(size=12)),
        yaxis=dict(title_font=dict(size=16), tickfont=dict(size=12), title_text=""),
        # Adjust facet title font size
        # annotations=[dict(font=dict(size=16)) for i in fig.layout.annotations] # This was for general annotations, need specific for facets
    )

    # Adjust facet titles (indicator names)
    fig.for_each_annotation(lambda a: a.update(text=a.text.split("=")[-1]))

    st.plotly_chart(fig, use_container_width=True)

    # Add download button
    html_bytes = fig.to_html(include_plotlyjs='cdn').encode('utf-8')
    st.download_button(
        "下載同業差異點狀圖(HTML)",
        data=html_bytes,
        file_name="peer_comparison_dot_plot.html",
        mime="text/html",
        key="dl_peer_dot_plot"
    )


    # 差異表（以 target 為基準）
    # base = bank[bank['公司'] == target][['公司'] + indicators].set_index('公司').iloc[0]
    # diffs = []
    # for p in peers:
    #     if p in df_sel['公司'].values:
    #         row = df_sel[df_sel['公司'] == p][['公司'] + indicators].set_index('公司').iloc[0]
    #         delta = (base - row).to_dict()
    #         delta['同業'] = p
    #         diffs.append(delta)
    # diff_df = pd.DataFrame(diffs).set_index('同業') if diffs else pd.DataFrame()

    # if not diff_df.empty:
    #     st.subheader("與同業差異（正值=目標較高，負值=目標較低）")
    #     st.dataframe(diff_df.style.background_gradient(cmap='RdYlGn'))
    #     csv = diff_df.to_csv().encode('utf-8-sig')
    #     st.download_button("下載差異表CSV", data=csv, file_name="peer_diff.csv", mime="text/csv")

    # return fig, diff_df

def group_radar(bank, indicators, peer_metric):
    if len(indicators) < 3:
        st.warning("雷達圖至少需要3個業務指標才能繪製。\n")
        return None

    if peer_metric == '平均數':
        public_agg = bank[bank['類別']=='公股'][indicators].mean()
        private_agg = bank[bank['類別']=='民營'][indicators].mean()
    else:
        public_agg = bank[bank['類別']=='公股'][indicators].median()
        private_agg = bank[bank['類別']=='民營'][indicators].median()

    mega_series = bank[bank['公司']=='兆豐金控'][indicators].iloc[0]

    plot_df = pd.DataFrame({
        '指標': indicators,
        '兆豐金控': mega_series.values,
        f'公股 {peer_metric}': public_agg.values,
        f'民營 {peer_metric}': private_agg.values
    })

    fig = go.Figure()
    
    # 兆豐金控
    fig.add_trace(go.Scatterpolar(
        r=plot_df['兆豐金控'], 
        theta=plot_df['指標'], 
        fill='toself', 
        name='兆豐金控', 
        opacity=0.7,
        line=dict(color='#F4B000', width=3),
        mode='lines+markers+text',
        text=plot_df['兆豐金控'].round(2),
        textfont=dict(
            color='red', 
            size=16,     
            family="Arial, sans-serif"
        ),
        marker=dict(color='#F4B000', size=8, symbol='diamond')
    ))
    
    # 其他分組
    for col, color, symbol in [(f'公股 {peer_metric}', '#2BA84A', 'circle'), (f'民營 {peer_metric}', '#E4572E', 'square')]:
        text_color = ''
        if '公股' in col:
            text_color = '#00008B'
        elif '民營' in col:
            text_color = '#8B0000'

        fig.add_trace(go.Scatterpolar(
            r=plot_df[col], 
            theta=plot_df['指標'], 
            fill='toself', 
            name=col, 
            opacity=0.3,
            line=dict(color=color, width=2),
            mode='lines+markers+text',
            text=plot_df[col].round(2),
            textfont=dict(
                color=text_color,
                size=14,
                family="Arial, sans-serif"
            ),
            marker=dict(size=8, symbol=symbol, color=color)
        ))

    fig.update_layout(
        title=f'公股/民營分組 vs 兆豐金控 雷達圖 ({peer_metric})', 
        polar=dict(
            radialaxis=dict(
                visible=True, 
                side='counterclockwise',
                tickfont=dict(color='black', size=12),
                linecolor='rgba(0,0,0,0.5)',
                gridcolor='rgba(0,0,0,0.2)'
            ),
            angularaxis=dict(
                linecolor='rgba(0,0,0,0.2)'
            )
        ),
        legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
        font=dict(size=14, color='black', family="Arial, sans-serif")
    )
    st.subheader("雷達圖指標實際數值")
    for index, row in plot_df.iterrows():
        indicator_name = row['指標']
        mega_value = row['兆豐金控']
        public_value = row[f'公股 {peer_metric}']
        private_value = row[f'民營 {peer_metric}']
        st.markdown(f"**{indicator_name}**: 兆豐金控: <span style=\"color:#F4B000\">{mega_value:.2f}</span>, 公股{peer_metric}: <span style=\"color:#2BA84A\">{public_value:.2f}</span>, 民營{peer_metric}: <span style=\"color:#E4572E\">{private_value:.2f}</span>", unsafe_allow_html=True)
    st.markdown("---")

    st.plotly_chart(fig, use_container_width=True, key='group_radar_chart')

    # 提供下載HTML
    html_bytes = fig.to_html(include_plotlyjs='cdn').encode('utf-8')
    st.download_button("下載雷達圖(HTML)", data=html_bytes, file_name="group_radar.html", mime="text/html", key="dl_group_radar")

    return fig

# =========================
# 介面
# =========================
# --- Global Data Preparation (moved from sidebar) ---
existing_cols = set(data_df.columns)
filtered_subsidiary_groups = {
    group: [col for col in metrics if col in existing_cols]
    for group, metrics in SUBSIDIARY_GROUPS.items()
}
filtered_subsidiary_groups = {k: v for k, v in filtered_subsidiary_groups.items() if v}

if not filtered_subsidiary_groups:
    st.error("檔案中未找到可識別的指標。\n")
    st.stop()

selected_sub_types = st.multiselect(
    "選擇子公司類型",
    options=list(filtered_subsidiary_groups.keys()),
    default=list(filtered_subsidiary_groups.keys()),
    key='selected_sub_types_key'
)

# Calculate all_inds based on selected_sub_types
all_inds = []
for s_type in selected_sub_types:
    all_inds.extend(filtered_subsidiary_groups[s_type])
all_inds = sorted(list(set(all_inds))) # Get unique and sort them

# Handle case where no subsidiary types are selected
if not selected_sub_types:
    st.warning("請至少選擇一個子公司類型。\n")
    st.stop()

defaul_inds = all_inds[:5] if len(all_inds) > 5 else all_inds

# Ensure at least one indicator is selected if all_inds is not empty
initial_default_selection = [ind for ind in ["ROE_202506", "存款餘額市占率", "外匯存款餘額市占率", "信託商品市占率", "公民營放款市占率"] if ind in all_inds]
if not initial_default_selection and all_inds:
    initial_default_selection = [all_inds[0]]

# Adjust the label for selected_inds multiselect
sub_type_label = "多種子公司類型" if len(selected_sub_types) > 1 else selected_sub_types[0]

col1, col2 = st.columns([1, 5])
with col1:
    st.image("logos/兆豐金.jfif", width=100)
with col2:
    st.title("兆豐金控業務視覺化與同業差異分析")
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

    #st.markdown("---")
    #st.header("資料設定")

    

    

# st.markdown("---") # Separator for the main content area
# st.header("業務指標設定")

# Calculate all_inds based on selected_sub_types (this part is already calculated in the sidebar)
# all_inds is available here.

# Handle case where no subsidiary types are selected (this part is already handled in the sidebar)
# This part is already handled in the sidebar.

# Ensure at least one indicator is selected if all_inds is not empty (this part is already handled in the sidebar)
# This part is already handled in the sidebar.

# Adjust the label for selected_inds multiselect (this part needs to be moved from sidebar)
sub_type_label = "多種子公司類型" if len(selected_sub_types) > 1 else selected_sub_types[0]
selected_inds = data_df.columns[1:-1].tolist() # Assign all business indicators, excluding '公司' and '類別'
# sort_by = st.selectbox("公司排序依據", options=['(不排序)'] + selected_inds, index=0, key="global_sort_by")
st.markdown("---") # Separator

# 分頁

#!!RENDER
def render_tab_同業比較全覽():
    # with tab2:
    st.subheader("同業金控比較分析")
    target = '兆豐金控' # 預設目標公司為兆豐金控
    peer_candidates = [c for c in data_df['公司'].tolist() if c != '兆豐金控']
    peers = st.multiselect("選擇1~4家同業比較", options=peer_candidates, default=['國泰金控', '富邦金控', '中信金控', '玉山金控'][:min(4,len(peer_candidates))], key='peers_selection_tab1')

    # Allow user to select a single indicator if multiple are chosen globally
    indicators_for_comparison = selected_inds
    if len(selected_inds) > 1:
        indicator_to_show = st.selectbox('請選擇一個指標以進行比較', selected_inds, key='tab2_indicator_select')
        indicators_for_comparison = [indicator_to_show]

    if not indicators_for_comparison:
        st.warning("請在左側「指標設定」中至少選擇一個指標。\n")
    else:
        # The peer_comparison function handles plotting and downloading
        peer_comparison(data_df, target, peers, indicators_for_comparison, category_filter)

    st.markdown("---")    


    st.subheader("同業比較全覽")
    if not selected_inds:
        st.warning("請在左側「指標設定」中至少選擇一個指標。\n")
    else:
        chart_type = st.radio("選擇圖表類型", ["Bar Chart", "雷達圖"], index=0, horizontal=True, key="tab1_chart_type")

        if chart_type == "Bar Chart":
            # Existing logic for bar chart (single indicator selection)
            if len(selected_inds) > 0:
                single_indicator_to_display = st.selectbox(
                    "選擇要顯示的業務指標",
                    options=data_df.columns[1:-1].tolist(), # Use all business indicators from data_df
                    key="tab1_single_indicator_select"
                )
                display_indicators = [single_indicator_to_display]
            else:
                display_indicators = []

            chart_sort_indicator = display_indicators[0] # Always sort by the single displayed indicator

            fig_overview = overview_charts(data_df, display_indicators, chart_sort_indicator, category_filter, peer_metric, sub_type_label)
            st.plotly_chart(fig_overview, use_container_width=True)
            html_bytes_overview = fig_overview.to_html(include_plotlyjs='cdn').encode('utf-8')

        else: # 雷達圖
            # New logic for radar chart (multi-select indicators)
            radar_chart_indicators = st.multiselect(
                "選擇雷達圖要呈現的業務指標 (至少3個)",
                options=data_df.columns[1:-1].tolist(), # Use all business indicators from data_df
                default=data_df.columns[1:-1].tolist()[:3], # Default to first 3 business indicators
                key="radar_chart_indicators_key"
            )

            if len(radar_chart_indicators) < 3:
                st.warning("雷達圖至少需要3個業務指標才能繪製，請選擇足夠的指標。\n")
            else:
                fig_radar_tab1 = create_radar_chart_for_tab1(data_df, radar_chart_indicators, category_filter, peer_metric)
                if fig_radar_tab1:
                    st.plotly_chart(fig_radar_tab1, use_container_width=True)
                    html_bytes_radar_tab1 = fig_radar_tab1.to_html(include_plotlyjs='cdn').encode('utf-8')
                    st.download_button("下載雷達圖(HTML)", data=html_bytes_radar_tab1, file_name="overview_radar.html", mime="text/html", key="dl_overview_radar")

        st.markdown("---")

        # --- Section 2: Ranking dataframe for comparison ---
        st.subheader("同業業務指標比較表")

        rank_df = add_ranking(data_df, selected_inds)
        show_cols = ['公司','類別'] + selected_inds
        
        def highlight_mega(row):
            if row['公司'] == '兆豐金控':
                return ['background-color: #F4B000'] * len(row)
            else:
                return [''] * len(row)

        display_df = rank_df[show_cols]
        
        peer_values = data_df[data_df['公司'] != '兆豐金控'][selected_inds]
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
        # col_heat1, col_heat2 = st.columns(2)

        # with col_heat1:
        #     # Business Metrics Heatmap (from original overview_charts)
        #     st.write("#### 業務指標熱力圖")
        #     data = data_df.copy()
        #     if category_filter in ['公股','民營']:
        #         data = data[data['類別'] == category_filter]
        #     heat_df = data.set_index('公司')[selected_inds]
        #     heat_df_filtered = heat_df.dropna(how='all')
            
        #     z = heat_df_filtered.values
        #     fig_heat = go.Figure(data=go.Heatmap(
        #         z=z,
        #         x=selected_inds,
        #         y=heat_df_filtered.index,
        #         colorscale='Blues',
        #         colorbar=dict(title='數值')
        #     ))
        #     fig_heat.update_layout(title="業務指標熱力圖", height=600)
        #     st.plotly_chart(fig_heat, use_container_width=True)
        #     html_bytes_heat = fig_heat.to_html(include_plotlyjs='cdn').encode('utf-8')
        #     st.download_button("下載業務指標熱力圖(HTML)", data=html_bytes_heat, file_name="heatmap.html", mime="text/html", key="dl_heatmap")

        # with col_heat2:
        #     # Ranking Heatmap (from original tab4)
        #     st.write("#### 名次熱力圖")
        #     sort_key_indicator = sort_by if sort_by in selected_inds else selected_inds[0]
        #     sort_col = f"{sort_key_indicator}_排名"
        #     # rank_df was already computed above
        #     sorted_rank_df = rank_df.sort_values(by=sort_col, ascending=True)

        #     heat = sorted_rank_df.set_index('公司')[[f"{c}_排名" for c in selected_inds]]
        #     heat = heat.replace(np.inf, np.nan)
            
        #     figr = go.Figure(data=go.Heatmap(
        #         z=z,
        #         x=heat.columns,
        #         y=heat.index,
        #         colorscale='Tealgrn',
        #         reversescale=True,
        #         colorbar=dict(title='名次 (1最好)'),
        #         text=heat.values,
        #         texttemplate="%{text:.0f}",
        #         textfont={"size":12}
        #     ))
        #     figr.update_layout(title=f"名次熱力圖 (依 {sort_key_indicator} 排序)", height=600)
        #     st.plotly_chart(figr, use_container_width=True)
        #     html_bytes_figr = figr.to_html(include_plotlyjs='cdn').encode('utf-8')
        #     st.download_button("下載名次熱力圖(HTML)", data=html_bytes_figr, file_name="rank_heatmap.html", mime="text/html", key="dl_rank_heatmap")
def render_tab_差異比較分析_雷達圖():
    # with tab3:
    st.subheader("公股/民營差異雷達圖(點擊下方圖示可手動個別呈現各類別雷達圖)")
    tab3_radar_indicators = st.multiselect(
        "選擇雷達圖要呈現的業務指標 (至少3個)",
        options=data_df.columns[1:-1].tolist(),
        default=["存款餘額市占率", "外匯存款餘額市占率", "信託商品市占率"],
        key="tab3_radar_indicators_key"
    )

    if len(tab3_radar_indicators) < 3:
        st.warning("雷達圖至少需要3個業務指標才能繪製，請選擇足夠的指標。\n")
    else:
        fig_radar = group_radar(data_df, tab3_radar_indicators, peer_metric)
def render_tab_業務指標熱力圖():
    # with tab4:
    st.subheader("業務指標熱力圖")

    # Allow user to select indicators for the heatmap
    heatmap_indicators = st.multiselect(
        "選擇熱力圖指標",
        options=all_inds,
        default=[ind for ind in ["存款餘額市占率", "外匯存款餘額市占率"] if ind in all_inds], # Default to specific indicators
        key='heatmap_indicators_selection'
    )

    if heatmap_indicators:
        # Apply category filter once
        filtered_data_df = data_df.copy()
        if category_filter in ['公股', '民營']:
            filtered_data_df = filtered_data_df[filtered_data_df['類別'] == category_filter]

        for indicator in heatmap_indicators:
            # Prepare data for treemap for the current indicator
            df_for_single_heatmap = filtered_data_df[['公司', indicator]].copy()
            
            # Rename the indicator column to '數值' for treemap
            df_for_single_heatmap.rename(columns={indicator: '數值'}, inplace=True)

            # Calculate Mega Bank's value and rank from the unfiltered data_df
            mega_value = data_df[data_df['公司'] == '兆豐金控'][indicator].iloc[0]
            
            # Calculate rank for all companies in the unfiltered data_df
            temp_rank_df = data_df[['公司', indicator]].copy()
            temp_rank_df['rank_col'] = temp_rank_df[indicator].rank(ascending=False, method='min')
            mega_rank = temp_rank_df[temp_rank_df['公司'] == '兆豐金控']['rank_col'].iloc[0]
            
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
                if d == '兆豐金控':
                    fig_treemap.data[0].marker.line.width = [6 if label == '兆豐金控' else 0 for label in fig_treemap.data[0].labels] # Even thicker border
                    fig_treemap.data[0].marker.line.color = ['white' if label == '兆豐金控' else 'rgba(0,0,0,0)' for label in fig_treemap.data[0].labels] # White color
                    break
            
            # Calculate peer metric value
            peer_values_for_indicator = data_df[data_df['公司'] != '兆豐金控'][indicator]
            if peer_metric == '平均數':
                peer_metric_value = peer_values_for_indicator.mean()
            else: # 中位數
                peer_metric_value = peer_values_for_indicator.median()

            # Calculate difference
            difference = mega_value - peer_metric_value
            
            st.markdown(f"<p style='font-size:24px; color:#FFFFFF;'><b>兆豐金控 {indicator}</b>: {mega_value:.2f}% (在 {num_companies} 家金控中排名第 {int(mega_rank)} 名) | 同業{peer_metric}: {peer_metric_value:.2f}% (差異: {difference:+.2f}%)</p>", unsafe_allow_html=True)

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
        st.info("請選擇至少一個業務指標以顯示熱力圖。\n")
        
def render_tab_歷年走勢圖():
    # with tab1:
    st.subheader("歷年走勢圖") # Renamed subheader
    
    # Load data from Dynamic.xlsx
    try:
        dynamic_roe_df = load_dynamic_data("Dynamic.xlsx")
        
        # Identify ROE columns (assuming they are all columns except the first one)
        roe_years = dynamic_roe_df.columns[1:].tolist()
        
        if not roe_years:
            st.warning("未在Dynamic.xlsx中找到ROE年度數據.\n")
        else:
            # Merge with data_df to get '類別' information
            # Ensure '公司' column is named correctly in dynamic_roe_df before merge
            dynamic_roe_df.rename(columns={dynamic_roe_df.columns[0]: '公司'}, inplace=True)
            
            # Select only '公司' and '類別' from data_df for merging
            company_category_df = data_df[['公司', '類別']]
            
            df_merged = pd.merge(dynamic_roe_df, company_category_df, on='公司', how='left')
            
            # Apply category filter if not "全部"
            if category_filter != "全部":
                df_merged = df_merged[df_merged['類別'] == category_filter]

            # Create 'plot_group' column
            df_merged['plot_group'] = np.where(df_merged['公司'] == '兆豐金控', '兆豐金控', df_merged['類別'])

            # Melt DataFrame for Plotly Express
            df_melted = df_merged.melt(id_vars=['公司', '類別', 'plot_group'], # Include new columns
                                            value_vars=roe_years, 
                                            var_name='年度', 
                                            value_name='ROE')

            # Add rank for dynamic sorting. Rank 1 is highest ROE.
            df_melted['rank'] = df_melted.groupby('年度')['ROE'].rank(method='first', ascending=False)
            df_melted['rank'] = df_melted['rank'].astype(int)

            # --- START OF LINE CHART BLOCK (MOVED UP) ---
            # st.subheader("歷年變化趨勢")
            # st.write("此圖表顯示各金控歷年變化趨勢。")

            all_companies_for_line_chart = df_melted['公司'].unique().tolist()
            default_selected_companies_for_line_chart = ["中信金", "玉山金", "華南金", "第一金", "兆豐金控"]

            selected_companies_for_line_chart = st.multiselect(
                "選擇檢視公司 (上限6家)",
                options=all_companies_for_line_chart,
                default=[comp for comp in default_selected_companies_for_line_chart if comp in all_companies_for_line_chart],
                key="line_chart_company_selector"
            )

            if not selected_companies_for_line_chart:
                st.warning("請至少選擇一家公司。\n")
                df_filtered_companies = pd.DataFrame() # Empty DataFrame to prevent errors
            elif len(selected_companies_for_line_chart) > 6:
                st.warning("最多只能選擇6家公司。請減少選擇。\n")
                df_filtered_companies = df_melted[df_melted['公司'].isin(selected_companies_for_line_chart[:6])] # Use first 6
            else:
                df_filtered_companies = df_melted[df_melted['公司'].isin(selected_companies_for_line_chart)]

            fig_roe_line = px.line(
                df_filtered_companies, # Use filtered DataFrame
                x='年度',
                y='ROE',
                color='公司',
                title='各金控歷年變化趨勢',
                labels={'ROE': 'ROE (%)', '年度': '年度'},
                hover_data={'ROE': ':.2f'}
            )

            # Highlight 兆豐金控 and add markers
            for trace in fig_roe_line.data:
                trace.mode = 'lines+markers' # Add markers
                trace.marker = dict(symbol='circle', size=8) # Set marker to thick circle
                if trace.name == '兆豐金控':
                    trace.line.width = 4 # Make line thicker
                    trace.line.color = '#F4B000' # Mocha Mousse color
                else:
                    trace.line.width = 2 # Default thickness for others
                    trace.line.dash = 'dash' # Dashed line for others

            st.plotly_chart(fig_roe_line, use_container_width=True)

            html_bytes_roe_line = fig_roe_line.to_html(include_plotlyjs='cdn').encode('utf-8')
            st.download_button(
                "下載歷年變化趨勢圖(HTML)",
                data=html_bytes_roe_line,
                file_name="roe_line_chart.html",
                mime="text/html",
                key="dl_roe_line_chart"
            )
            # --- END OF LINE CHART BLOCK ---

            # Define a consistent color map for all companies
            all_companies = df_merged['公司'].unique()
            company_colors = {
                '兆豐金控': '#F4B000', '富邦金控': '#D64550', '國泰金控': '#E4572E',
                '中信金控': '#FF7C43', '玉山金控': '#FFA600', '元大金控': '#9A348E',
                '台新金控': '#6A057F', '新光金控': '#F08080', '開發金控': '#C70039',
                '永豐金控': '#900C3F', '臺灣金控': '#127C90', '合庫金控': '#2BA84A',
                '第一金控': '#005f73', '華南金控': '#0a9396'
            }
            color_sequence = px.colors.qualitative.Plotly
            color_idx = 0
            for company in all_companies:
                if company not in company_colors:
                    company_colors[company] = color_sequence[color_idx % len(color_sequence)]
                    color_idx += 1

                        # --- START OF BAR CHART BLOCK (REBUILT WITH GO.FIGURE) ---
            st.subheader("歷年走勢圖 (動態排序)")
            
            # Re-add the animation speed slider, as it's part of the dynamic chart UI
            animation_speed = st.slider("調整動畫速度 (毫秒)", min_value=1000, max_value=2000, value=1500, step=100)
            bar_gap_roe = st.slider('調整Bar條間距 (歷年走勢圖)', 0.0, 0.5, 0.1, 0.05)

            # Sort data by year to ensure correct frame order
            years = sorted(df_melted['年度'].unique())
            
            # --- Create the figure with manual frames ---
            fig = go.Figure()

            # --- Add data and layout for the first year (initial view) ---
            initial_year = years[0]
            initial_df = df_melted[df_melted['年度'] == initial_year].sort_values('rank', ascending=True)

            # Define colors based on category for the initial view
            initial_colors = []
            for _, row in initial_df.iterrows():
                if row['公司'] == '兆豐金控':
                    initial_colors.append('#F4B000')
                elif row['類別'] == '公股':
                    initial_colors.append('#127C90')
                else:
                    initial_colors.append('#D64550')

            fig.add_trace(go.Bar(
                x=initial_df['ROE'],
                y=initial_df['公司'],
                orientation='h',
                text=initial_df['公司'],
                textposition='inside',
                insidetextanchor='middle',
                marker=dict(
                    color=initial_colors,
                    opacity=[0.5 if c != '兆豐金控' else 1.0 for c in initial_df['公司']]
                ),
                textfont={'color':'white', 'size':16}
            ))

            # --- Create a frame for each year ---
            frames = []
            for year in years:
                frame_df = df_melted[df_melted['年度'] == year].sort_values('rank', ascending=True)
                
                # Create annotations for this frame
                annotations = []
                for _, row in frame_df.iterrows():
                    annotations.append(go.layout.Annotation(
                        text=f"<b>#{row['rank']}</b>",
                        align='right', showarrow=False, xref='paper', yref='y',
                        x=0, y=row['公司'], xanchor='right',
                        font=dict(color="#4169E1", size=16)
                    ))
                
                # Add year annotation to the bottom right
                annotations.append(go.layout.Annotation(
                    text=str(year),
                    align='right',
                    showarrow=False,
                    xref='paper',
                    yref='paper',
                    x=0.95,
                    y=0.05,
                    xanchor='right',
                    yanchor='bottom',
                    font=dict(size=80, color="#4169E1")
                ))
                
                # Define colors based on category for each frame
                frame_colors = []
                for _, row in frame_df.iterrows():
                    if row['公司'] == '兆豐金控':
                        frame_colors.append('#F4B000')
                    elif row['類別'] == '公股':
                        frame_colors.append('#127C90')
                    else:
                        frame_colors.append('#D64550')

                frames.append(go.Frame(
                    name=str(year),
                    data=[go.Bar(
                        x=frame_df['ROE'],
                        y=frame_df['公司'],
                        orientation='h',
                        text=frame_df['公司'],
                        textposition='inside',
                        insidetextanchor='middle',
                        marker=dict(
                            color=frame_colors,
                            opacity=[0.5 if c != '兆豐金控' else 1.0 for c in frame_df['公司']]
                        ),
                        textfont={'color':'white', 'size':16}
                    )],
                    layout=go.Layout(annotations=annotations)
                ))
            
            fig.frames = frames

            # --- Create and configure the slider ---
            sliders = [dict(
                active=0,
                currentvalue={"prefix": "年度: ", "font": {"size": 25}},
                pad={"t": 50},
                steps=[dict(
                    label=str(year),
                    method="animate",
                    args=[[str(year)], dict(
                        mode="immediate",
                        frame=dict(duration=animation_speed, redraw=True),
                        transition=dict(duration=animation_speed)
                    )]
                ) for year in years]
            )]

            # --- Update the main layout ---
            fig.update_layout(
                title='歷年走勢BAR圖 (動態排序)',
                height=1000,
                margin=dict(l=100),
                bargap=bar_gap_roe,
                yaxis=dict(autorange="reversed", showticklabels=False),
                updatemenus=[dict(
                    type="buttons",
                    showactive=False,
                    buttons=[dict(
                        label="Play",
                        method="animate",
                        args=[None, dict(
                            frame=dict(duration=animation_speed, redraw=True),
                            fromcurrent=True,
                            transition=dict(duration=animation_speed, easing="linear")
                        )]
                    ), dict(
                        label="Pause",
                        method="animate",
                        args=[[None], dict(
                            mode="immediate"
                        )]
                    )]
                )],
                sliders=sliders
            )
            
            # Set initial annotations
            fig.update_layout(annotations=frames[0].layout.annotations)

            st.plotly_chart(fig, use_container_width=True)
            
            # Add download button back
            html_bytes_roe_dynamic = fig.to_html(include_plotlyjs='cdn').encode('utf-8')
            st.download_button(
                "下載歷年走勢圖(HTML)",
                data=html_bytes_roe_dynamic,
                file_name="roe_dynamic_chart.html",
                mime="text/html",
                key="dl_roe_dynamic_chart"
            )
            # --- END OF REBUILT BAR CHART BLOCK ---

    except FileNotFoundError:
        st.error("找不到Dynamic.xlsx檔案，請確認檔案是否存在於相同目錄.\n")
    except Exception as e:
        st.error(f"讀取或處理Dynamic.xlsx時發生錯誤: {e}\n")
        
def render_tab_自選指標趨勢():    
    # with tab5:
    st.subheader("自選指標趨勢圖")

    # Load data from Dynamic2.xlsx
    dynamic2_data, sheet_names = load_dynamic2_data("Dynamic2.xlsx")

    if not sheet_names:
        st.warning("找不到 'Dynamic2.xlsx' 或該檔案中沒有任何工作表。")
    else:
        # UI Controls
        selected_sheet = st.selectbox("請選擇指標 (Sheet):", options=sheet_names, key="dynamic2_sheet_selector")
        
        if selected_sheet and dynamic2_data:
            df_indicator_raw = dynamic2_data[selected_sheet]
            
            # Merge with data_df to get '類別' information
            company_category_df = data_df[['公司', '類別']]
            df_indicator = pd.merge(df_indicator_raw, company_category_df, on='公司', how='left')

            all_companies = df_indicator['公司'].unique().tolist()
            
            # Default selection logic
            if selected_sheet == "銀行_ROE":
                default_selection = ['兆豐銀行', '中信銀行', '玉山銀行', '富邦銀行', '第一銀行', '合庫銀行']
            else:
                default_selection = ['兆豐金控']
                other_default_companies = ['富邦金控', '中信金控', '國泰金控', '玉山金控', '第一金控']
                default_selection.extend([c for c in other_default_companies if c in all_companies])

            # Ensure unique and valid companies in default selection
            default_selection = [c for c in default_selection if c in all_companies]

            selected_companies = st.multiselect(
                "請選擇公司 (最多6家):",
                options=all_companies,
                default=default_selection,
                key="dynamic2_company_selector"
            )

            if not selected_companies:
                st.info("請至少選擇一家公司。")
            
            # Truncate if more than 6 are selected
            if len(selected_companies) > 6:
                st.warning("最多只能選擇6家公司，將只顯示前6家。")
                selected_companies = selected_companies[:6]
            
            if selected_companies:
                # Filter data and plot
                df_plot = df_indicator[df_indicator['公司'].isin(selected_companies)]
                
                # Melt dataframe
                id_vars = ['公司', '類別']
                
                if 'ROE' in selected_sheet.upper(): # Check for ROE in sheet name (case-insensitive)
                    value_vars = [col for col in df_plot.columns if col not in id_vars]
                else:
                    # Filter for columns containing "市占率"
                    value_vars = [col for col in df_plot.columns if "市占率" in col and col not in id_vars]
                    if not value_vars: # Fallback if no "市占率" columns are found
                        st.warning(f"在 '{selected_sheet}' 工作表中未找到任何 '市占率' 相關指標。將顯示所有可用指標。")
                        value_vars = [col for col in df_plot.columns if col not in id_vars]

                df_melted = df_plot.melt(
                    id_vars=id_vars,
                    value_vars=value_vars,
                    var_name='年度',
                    value_name=selected_sheet
                )

                # Create color and line style maps
                color_map = {}
                line_dash_map = {}

                if selected_sheet == "銀行_ROE":
                    public_banks = ['華南銀行', '第一銀行', '合庫銀行', '臺灣銀行']
                    for company in selected_companies:
                        if company == '兆豐銀行':
                            color_map[company] = '#F4B000'
                            line_dash_map[company] = 'solid'
                        elif company in public_banks:
                            color_map[company] = '#2BA84A' # Public color
                            line_dash_map[company] = 'dash'
                        else:
                            color_map[company] = '#E4572E' # Private color
                            line_dash_map[company] = 'dash'
                else:
                    for company in selected_companies:
                        if company in ['兆豐金控', '兆豐銀行']:
                            color_map[company] = '#F4B000'
                            line_dash_map[company] = 'solid'
                        else:
                            try:
                                category = df_indicator[df_indicator['公司'] == company]['類別'].iloc[0]
                                if category == '民營':
                                    color_map[company] = '#E4572E'
                                elif category == '公股':
                                    color_map[company] = '#2BA84A'
                                else:
                                    color_map[company] = '#808080' # Fallback
                                line_dash_map[company] = 'dash'
                            except (IndexError, KeyError):
                                color_map[company] = '#808080' # Fallback
                                line_dash_map[company] = 'dash'

                # Create line chart
                fig_dynamic2 = px.line(
                    df_melted,
                    x='年度',
                    y=selected_sheet,
                    color='公司',
                    line_dash='公司', # Use company name for dash mapping
                    title=f'"'+selected_sheet+'" 歷年走勢',
                    markers=True,
                    color_discrete_map=color_map,
                    line_dash_map=line_dash_map
                )
                
                # Set line widths
                for trace in fig_dynamic2.data:
                    if trace.name in ['兆豐金控', '兆豐銀行']:
                        trace.line.width = 4
                    else:
                        trace.line.width = 2
                
                # Set Y-axis to start from 0
                fig_dynamic2.update_yaxes(rangemode="tozero")

                st.plotly_chart(fig_dynamic2, use_container_width=True)
                    
# !!!!!!!PERPLEXITY TAB6

# 於現有 tabs 定義處加入 TAB6
# tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["歷年走勢圖", "同業比較全覽", "差異比較分析(雷達圖)", "業務指標熱力圖", "自選指標趨勢"])
def render_tab_戰情室():
# ===== Tab6 複數歷年折線圖（6x2 共 12 張） =====
    # with tab6:
    st.subheader("各業務市占率總覽")

    # 動態載入 Dynamic2.xlsx 全部工作表
    # 優先讀同資料夾 Dynamic2.xlsx，如需指定完整路徑可替換參數
    data_sheets, sheet_names = load_dynamic2_data("Dynamic2.xlsx")

    if data_sheets is None or not sheet_names:
        st.warning("未找到 Dynamic2.xlsx 或無可用工作表，請確認檔案位置與內容。")
        st.stop()

    # 前 9 張圖固定標題（依需求命名）
    titles_first_9 = [
        "金控ROE",
        "銀行ROE",
        "存款餘額市占率",
        "外匯存款餘額市占率",
        "公民營事業放款市占率",
        "中小企業放款市占率",
        "消金貸款市占率",
        "信託商品市占率",
        "衍生性金融商品餘額市占率",
    ]

    # 建立 12 張圖的設定（索引 0~11）
    chart_configs = []

    # 0-8：對應前 9 個工作表
    for i in range(9):
        chart_configs.append({
            "title": titles_first_9[i],
            "sheet_index": i,
            "placeholder": False
        })

    # 9：佔位
    chart_configs.append({
        "title": "佔位空間待補",
        "sheet_index": None,
        "placeholder": True
    })

    # 10、11：使用工作表索引 10 與 11（若 11 不存在則顯示佔位），標題前加「(證券)」
    # 第 11 張圖（索引 10）對應 Excel 第 11 張（索引 10）
    if len(sheet_names) > 9: # Corrected: check for index 9
        sec_name_10 = sheet_names[9] # Corrected: access specific sheet name by index
        sec_title_10 = f"(證券){sec_name_10.split('_', 1)[-1]}市占率" if "ROE" not in sec_name_10 else f"(證券){sec_name_10.split('_', 1)[-1]}"
        chart_configs.append({
            "title": sec_title_10,
            "sheet_index": 9, # Corrected: use index 9 for the 10th sheet
            "placeholder": False
        })
    else:
        chart_configs.append({
            "title": "(證券)資料待補",
            "sheet_index": None,
            "placeholder": True
        })

    # 第 12 張圖（索引 11）對應 Excel 第 12 張（索引 11）
    if len(sheet_names) > 10: # Corrected: check for index 10
        sec_name_11 = sheet_names[10] # Corrected: access specific sheet name by index
        sec_title_11 = f"(證券){sec_name_11.split('_', 1)[-1]}市占率" if "ROE" not in sec_name_11 else f"(證券){sec_name_11.split('_', 1)[-1]}"
        chart_configs.append({
            "title": sec_title_11,
            "sheet_index": 10, # Corrected: use index 10 for the 11th sheet
            "placeholder": False
        })
    else:
        chart_configs.append({
            "title": "(證券)資料待補",
            "sheet_index": None,
            "placeholder": True
        })

    # 預設公司（銀行/金控）
    default_banks = ["中信銀行", "玉山銀行", "富邦銀行", "第一銀行", "兆豐銀行", "合庫銀行"]
    default_fhc   = ["中信金控", "玉山金控", "富邦金控", "第一金控", "兆豐金控", "合庫金控"]

    public_banks = {"合庫銀行", "第一銀行", "華南銀行", "臺灣銀行"}
    public_fhc   = {"華南金控", "第一金控", "合庫金控", "臺灣金控"}  # 兆豐金控另行以亮黃粗實線標示

    def get_time_columns(df, mode):
        # mode: "roe" or "share"
        cols = [c for c in df.columns if c != "公司"]
        if mode == "roe":
            # 全數值欄
            return [c for c in cols if pd.api.types.is_numeric_dtype(df[c])]
        else:
            # 僅取含「市占率」欄
            return [c for c in cols if ("市占率" in str(c)) and pd.api.types.is_numeric_dtype(df[c])]

    def render_line_chart(df, title, time_cols, entity_type, key_suffix, defaults):
        # entity_type: "bank" or "fhc"
        options = df["公司"].dropna().unique().tolist()
        default_vals = [c for c in defaults if c in options]

        selector_key = f"tab6_sel_{key_suffix}"
        selected = st.multiselect(f"{title}－選擇檢視公司（上限6家）", options=options, default=default_vals, key=selector_key)

        if not selected:
            st.info("請至少選擇一家公司。")
            return

        if len(selected) > 6:
            st.warning("最多只能選擇 6 家，已自動採前 6 家。")
            selected = selected[:6]

        df_sel = df[df["公司"].isin(selected)][["公司"] + time_cols].copy()
        long_df = df_sel.melt(id_vars="公司", var_name="期間", value_name="值")

        fig = go.Figure()
        for comp in selected:
            sub = long_df[long_df["公司"] == comp]
            # 樣式規則
            color = "#E4572E"   # 民營紅
            dash  = "dash"
            width = 2

            if entity_type == "bank":
                if comp == "兆豐銀行":
                    color, dash, width = "#F4B000", "solid", 4
                elif comp in public_banks:
                    color, dash, width = "#2BA84A", "dash", 2
                else:
                    color, dash, width = "#E4572E", "dash", 2
            else:  # fhc
                if comp == "兆豐金控":
                    color, dash, width = "#F4B000", "solid", 4
                elif comp in public_fhc:
                    color, dash, width = "#2BA84A", "dash", 2
                else:
                    color, dash, width = "#E4572E", "dash", 2

            fig.add_trace(go.Scatter(
                x=sub["期間"], y=sub["值"],
                mode="lines+markers",
                name=comp,
                line=dict(color=color, width=width, dash=dash),
                marker=dict(size=8)
            ))

            # Add company name to the right end of each line (Point 2)
            if not sub.empty:
                last_data_point = sub.iloc[-1]
                fig.add_annotation(
                    x=last_data_point["期間"],
                    y=last_data_point["值"],
                    text=comp,
                    showarrow=False,
                    xanchor="left",
                    xshift=10, # Shift text to the right
                    font=dict(color=color, size=10) # Use line color for text
                )

        fig.update_layout(
            title=title,
            xaxis_title="年度", # Changed from "期間" to "年度" for clarity (Point 3)
            yaxis_title="", # Removed "數值" (Point 1)
            legend_title="公司",
            height=400,
            margin=dict(l=40, r=20, t=60, b=40),
            xaxis=dict(type='category') # Ensure X-axis is treated as categories (Point 3)
        )
        st.plotly_chart(fig, use_container_width=True, key=f"chart_{key_suffix}")

        # 下載 HTML
        html_bytes = fig.to_html(include_plotlyjs='cdn').encode('utf-8')
        st.download_button(
            f"下載 {title} (HTML)",
            data=html_bytes,
            file_name=f"{key_suffix}.html",
            mime="text/html",
            key=f"dl_{key_suffix}"
        )

    # 依 6x2 佈局輸出 12 張圖
    for row in range(6):
        col_left, col_right = st.columns(2)
        for col_idx, col in enumerate([col_left, col_right]):
            chart_idx = row * 2 + col_idx
            cfg = chart_configs[chart_idx]

            with col:
                st.markdown(f"### {cfg['title']}")

                if cfg["placeholder"]:
                    st.info("佔位空間待補")
                    continue

                si = cfg["sheet_index"]
                if si is None or si >= len(sheet_names):
                    st.info("資料待補（未找到對應工作表）")
                    continue

                sname = sheet_names[si]
                df = data_sheets.get(sname)
                if df is None or df.empty:
                    st.info("該工作表無資料")
                    continue

                # 判斷繪圖模式與預設公司
                if cfg["title"].startswith("金控ROE"):
                    time_cols = get_time_columns(df, "roe")[:-1]
                    entity_type = "fhc"
                    defaults = default_fhc
                elif cfg["title"].startswith("銀行ROE"):
                    time_cols = get_time_columns(df, "roe")[:-1]
                    entity_type = "bank"
                    defaults = default_banks
                else:
                    time_cols = get_time_columns(df, "share")
                    # 其餘圖以金控層級預設
                    entity_type = "fhc"
                    defaults = default_fhc

                if not time_cols:
                    st.info("未偵測到可繪製的『市占率』欄位或數值欄位。")
                    continue

                render_line_chart(
                    df=df,
                    title=cfg["title"],
                    time_cols=time_cols,
                    entity_type=entity_type,
                    key_suffix=f"tab6_{chart_idx}",
                    defaults=defaults
                )
                
# !!!3) 建立「標籤 -> 渲染函式」對照表（鍵值要與 UI 顯示的分頁文字一致）
TAB_RENDERERS = {
    "歷年走勢圖": render_tab_歷年走勢圖,
    "同業比較全覽": render_tab_同業比較全覽,
    "差異比較分析(雷達圖)": render_tab_差異比較分析_雷達圖,
    "業務指標熱力圖": render_tab_業務指標熱力圖,
    "自選指標趨勢": render_tab_自選指標趨勢,
    "戰情室": render_tab_戰情室,
}

# 4) 建立預設順序（只在第一次進入時設定）
if "tab_order" not in st.session_state:
    st.session_state["tab_order"] = list(TAB_RENDERERS.keys())

# 5) 提供「拖曳排序」UI（可放在主畫面頂端或側欄 expander 中）
with st.expander("拖曳調整分頁順序", expanded=False):
    # streamlit-sortables 僅支援 list[str] 與複數容器模式，這裡採最簡單的單清單模式
    new_order = sort_items(st.session_state["tab_order"].copy(), key="tab_sorter")
    if new_order and new_order != st.session_state["tab_order"]:
        st.session_state["tab_order"] = new_order
        st.toast("已更新分頁順序")
        st.rerun()

# 6) 依目前順序動態建立 tabs，並渲染相對應內容
tabs = st.tabs(st.session_state["tab_order"])
for label, tab in zip(st.session_state["tab_order"], tabs):
    with tab:
        TAB_RENDERERS[label]()  # 呼叫對應分頁的渲染函式