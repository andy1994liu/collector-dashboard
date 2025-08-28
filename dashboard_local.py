import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import io

# --- 頁面設定 (Page Configuration) ---
st.set_page_config(
    page_title="催收人員行為儀表板",
    page_icon="📊",
    layout="wide"
)

# --- 核心功能函式 (Core Functions) ---

def get_gdrive_download_url(share_link):
    """將 Google Drive 分享連結轉換為直接下載連結"""
    file_id = None
    if 'spreadsheets/d/' in share_link:
        file_id = share_link.split('/d/')[1].split('/')[0]
        return f'https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx'
    elif 'file/d/' in share_link:
        file_id = share_link.split('/d/')[1].split('/')[0]
        return f'https://drive.google.com/uc?export=download&id={file_id}'
    
    st.warning(f"提供的連結格式不正確，請確認是有效的 Google Drive 檔案分享連結: {share_link}")
    return None

@st.cache_data # 使用快取加速資料讀取
def load_data_from_gdrive(visit_url, group_url):
    """
    從 Google Drive 分享連結讀取並處理指定的 Excel 檔案。
    - visit_url: KH_DM_COLL_REPOSN_VISIT.xlsx 的分享連結
    - group_url: 分組名單.xlsx 的分享連結
    返回: 處理完成的 visit_logs (DataFrame) 和 groups_info (DataFrame)
    """
    visit_download_url = get_gdrive_download_url(visit_url)
    group_download_url = get_gdrive_download_url(group_url)

    if not visit_download_url or not group_download_url:
        return None, None

    try:
        groups_df = pd.read_excel(group_download_url)
        groups_df.columns = groups_df.columns.str.strip()
        groups_df['ID'] = groups_df['ID'].astype(str).str.strip()
        
        if 'Collector' in groups_df.columns and 'Agent Name' not in groups_df.columns:
            groups_df.rename(columns={'Collector': 'Agent Name'}, inplace=True)

        collectors_info = groups_df.set_index('ID')['Agent Name'].to_dict()

        visit_logs_df = pd.read_excel(visit_download_url, header=1)
        visit_logs_df.columns = visit_logs_df.columns.str.strip()

        required_cols = ['Collector', 'Create Time', 'Neg Pos Unit', 'Aging', 'Contact Summary', 'Customer Name']
        if not all(col in visit_logs_df.columns for col in required_cols):
            st.error("外訪紀錄檔缺少必要的欄位，請確認檔案包含：'Collector', 'Create Time', 'Neg Pos Unit', 'Aging', 'Contact Summary', 'Customer Name'")
            return None, None
            
        visit_logs_df = visit_logs_df[required_cols].copy()

        visit_logs_df.dropna(subset=['Create Time', 'Collector', 'Aging'], inplace=True)
        visit_logs_df['Collector'] = visit_logs_df['Collector'].astype(str).str.strip()
        
        visit_logs_df['Collector ID'] = visit_logs_df['Collector'].str.split('-').str[0].str.strip()
        
        valid_collector_ids = groups_df['ID'].unique()
        visit_logs_df = visit_logs_df[visit_logs_df['Collector ID'].isin(valid_collector_ids)]

        if 'Neg Pos Unit' in visit_logs_df.columns:
            visit_logs_df['Neg Pos Unit'] = visit_logs_df['Neg Pos Unit'].astype(str)
            visit_logs_df = visit_logs_df[visit_logs_df['Neg Pos Unit'].str.strip() != 'CancelRepossession']

        visit_logs_df['Create Time'] = pd.to_datetime(visit_logs_df['Create Time'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
        visit_logs_df.dropna(subset=['Create Time'], inplace=True)
        visit_logs_df['Date'] = visit_logs_df['Create Time'].dt.date

        visit_logs_df['Collector Name'] = visit_logs_df['Collector ID'].map(collectors_info)
        group_map = groups_df.set_index('ID')['Group'].to_dict()
        visit_logs_df['Group'] = visit_logs_df['Collector ID'].map(group_map)
        
        visit_logs_df['Aging'] = visit_logs_df['Aging'].astype(str).str.strip().str.upper()
        
        return visit_logs_df, groups_df

    except Exception as e:
        st.error(f"讀取 Google Drive 檔案時發生錯誤，請檢查連結權限或檔案格式。錯誤訊息：{e}")
        return None, None

def get_aging_colors(aging_status):
    """根據 Aging 狀態返回邊框和背景的顏色 (全新色系)"""
    default_colors = {'border': "#6B7280", 'bg': "rgba(107, 114, 128, 0.1)"} # 灰色
    if not isinstance(aging_status, str):
        return default_colors

    aging_status = aging_status.strip().upper()
    
    # --- 修改：採用視覺區分度更高的全新色系 ---
    color_map = {
        'M2': {'border': "#FBBF24", 'bg': "rgba(251, 191, 36, 0.15)"},  # 黃色
        'M3': {'border': "#F97316", 'bg': "rgba(249, 115, 22, 0.15)"},  # 橘色
        'M4': {'border': "#EF4444", 'bg': "rgba(239, 68, 68, 0.15)"},   # 紅色
        'M5': {'border': "#BE123C", 'bg': "rgba(190, 18, 60, 0.15)"},   # 深紅色/玫瑰紅
    }
    
    if 'M6' in aging_status:
        return {'border': "#86198F", 'bg': "rgba(134, 25, 143, 0.15)"} # 紫色 (最高警示)
    
    return color_map.get(aging_status, default_colors)

def display_case_card(row):
    """顯示單一案件卡片"""
    colors = get_aging_colors(row['Aging'])
    border_color = colors['border']
    bg_color = colors['bg']
    
    with st.container():
        st.markdown(f"""
        <div style="border-left: 5px solid {border_color}; border-radius: 5px; padding: 10px; background-color: {bg_color}; margin-bottom: 8px; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05);">
            <p style="font-weight: bold; color: #111827; margin: 0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="{row['Customer Name']}">
                {row['Customer Name'] if pd.notna(row['Customer Name']) else 'N/A'}
            </p>
            <div style="display: flex; justify-content: space-between; font-size: 0.8em; color: #4B5563; margin-top: 4px;">
                <span style="font-weight: bold; color: {border_color};">Aging: {row['Aging'] if pd.notna(row['Aging']) else 'N/A'}</span>
                <span title="{row['Collector Name'] if pd.notna(row['Collector Name']) else row['Collector']}">
                    {row['Collector Name'].split('-')[-1] if pd.notna(row['Collector Name']) else row['Collector']}
                </span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        with st.expander("View Summary"):
            st.markdown(f"**紀錄時間 (Time):** `{row['Create Time'].strftime('%Y-%m-%d %H:%M:%S')}`")
            st.markdown(f"**案件資訊 (Case Info):** `Aging: {row['Aging']} | Unit: {row['Neg Pos Unit']}`")
            st.markdown("**催收摘要 (Contact Summary):**")
            st.info(f"{row['Contact Summary'] if pd.notna(row['Contact Summary']) else '無摘要資訊'}")

def get_heatmap_color(count, max_count):
    """計算月曆熱力圖顏色"""
    if count == 0:
        return "#F3F4F6"
    normalized = (count / (max_count or 1)) * 0.9 + 0.1 
    lightness = 95 - (normalized * 55)
    return f"hsl(220, 80%, {lightness}%)"

def switch_to_week_view(date_to_view):
    """切換至週視圖的回呼函式"""
    st.session_state.current_date = date_to_view
    st.session_state.view_mode = '週'


# --- 主應用程式介面 (Main App Interface) ---
st.title("📊 催收人員行為儀表板")
st.caption("Collector Activity Dashboard (Python + Streamlit)")

# --- 側邊欄 (Sidebar) ---
with st.sidebar:
    st.header("⚙️ 資料來源設定")
    st.markdown("請提供 Google Drive 中 `KH_DM_COLL_REPOSN_VISIT.xlsx` 和 `分組名單.xlsx` 的分享連結。")

    # 使用 st.secrets 儲存連結以方便部署，如果不存在則使用 text_input
    default_visit_url = st.secrets.get("GDRIVE_VISIT_URL", "")
    default_group_url = st.secrets.get("GDRIVE_GROUP_URL", "")

    visit_log_url = st.text_input(
        "外訪紀錄 (VISIT) 分享連結", 
        value=default_visit_url,
        help="將 Google Drive 檔案權限設定為「知道連結的任何人」"
    )
    group_list_url = st.text_input(
        "分組名單 (GROUP) 分享連結", 
        value=default_group_url,
        help="將 Google Drive 檔案權限設定為「知道連結的任何人」"
    )

    st.header("篩選條件")
    
    if visit_log_url and group_list_url:
        visit_logs, groups_info = load_data_from_gdrive(visit_log_url, group_list_url)
    else:
        visit_logs, groups_info = None, None

    if visit_logs is not None and groups_info is not None:
        group_list = ['所有團隊'] + sorted(groups_info['Group'].unique().tolist())
        selected_group = st.selectbox('選擇團隊 (Group)', group_list)

        if selected_group == '所有團隊':
            collectors_in_group = ['所有催收員'] + sorted(groups_info['Agent Name'].unique().tolist())
        else:
            collector_ids_in_group = groups_info[groups_info['Group'] == selected_group]['ID'].tolist()
            filtered_collectors = visit_logs[visit_logs['Collector ID'].isin(collector_ids_in_group)]
            collectors_in_group = ['所有催收員'] + sorted(filtered_collectors['Collector Name'].dropna().unique().tolist())
            
        selected_collector_name = st.selectbox('選擇催收員 (Collector)', collectors_in_group)
    else:
        st.warning("請在上方提供有效的 Google Drive 分享連結以載入資料。")

# --- 主面板 (Main Panel) ---
if visit_logs is not None and groups_info is not None:
    base_filtered_data = visit_logs.copy()
    if selected_group != '所有團隊':
        base_filtered_data = base_filtered_data[base_filtered_data['Group'] == selected_group]
    if selected_collector_name != '所有催收員':
        base_filtered_data = base_filtered_data[base_filtered_data['Collector Name'] == selected_collector_name]

    tab1, tab2, tab3 = st.tabs(["🗓️ 行為追蹤", "📈 個人績效", "📊 行為模式比較"])

    # --- 行為追蹤 Tab ---
    with tab1:
        st.radio("切換視圖", ('週', '月'), key='view_mode', horizontal=True, label_visibility="collapsed")
        
        if 'current_date' not in st.session_state:
            st.session_state.current_date = datetime.now().date()

        col1, col2, col3, col4, col5 = st.columns([1, 1, 3, 1, 1])
        
        if col1.button('◀️ 上一' + ('週' if st.session_state.view_mode == '週' else '月')):
            if st.session_state.view_mode == '週':
                st.session_state.current_date -= timedelta(weeks=1)
            else:
                current_dt = datetime.combine(st.session_state.current_date, datetime.min.time())
                first_day_of_current_month = current_dt.replace(day=1)
                last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
                st.session_state.current_date = last_day_of_previous_month.date()
            st.rerun()

        if col2.button('今天'):
            st.session_state.current_date = datetime.now().date()
            st.rerun()

        if col5.button('下一' + ('週' if st.session_state.view_mode == '週' else '月') + ' ▶️'):
            if st.session_state.view_mode == '週':
                st.session_state.current_date += timedelta(weeks=1)
            else:
                current_dt = datetime.combine(st.session_state.current_date, datetime.min.time())
                first_day_of_current_month = current_dt.replace(day=1)
                next_month_first_day = (first_day_of_current_month.replace(day=28) + timedelta(days=4)).replace(day=1)
                st.session_state.current_date = next_month_first_day.date()
            st.rerun()

        if st.session_state.view_mode == '週':
            start_of_week = st.session_state.current_date - timedelta(days=st.session_state.current_date.weekday())
            end_of_week = start_of_week + timedelta(days=6)
            col3.subheader(f"{start_of_week.strftime('%Y/%m/%d')} - {end_of_week.strftime('%Y/%m/%d')}")
            
            days_of_week = [start_of_week + timedelta(days=i) for i in range(7)]
            day_cols = st.columns(7)
            
            for i, day in enumerate(days_of_week):
                with day_cols[i]:
                    is_today = (day == datetime.now().date())
                    header_html = f"""
                    <p style='text-align: center; font-weight: bold; background-color: {"#1E40AF" if is_today else "#D1D5DB"}; color: {"white" if is_today else "black"}; padding: 5px; border-radius: 5px;'>
                        {day.strftime('%A')}<br>{day.strftime('%m/%d')}
                    </p>"""
                    st.markdown(header_html, unsafe_allow_html=True)
                    
                    day_data = base_filtered_data[base_filtered_data['Date'] == day]
                    st.metric(label="案件數", value=len(day_data))
                    st.markdown("---")
                    
                    if not day_data.empty:
                        day_data['AgingSort'] = day_data['Aging'].str.extract('(\d+)').fillna(0).astype(int)
                        for _, row in day_data.sort_values('AgingSort', ascending=False).iterrows():
                            display_case_card(row)
                    else:
                        st.caption("無紀錄")

        if st.session_state.view_mode == '月':
            current_month_date = st.session_state.current_date
            col3.subheader(f"{current_month_date.year}年 {current_month_date.month}月")

            daily_counts = base_filtered_data[(base_filtered_data['Create Time'].dt.year == current_month_date.year) & (base_filtered_data['Create Time'].dt.month == current_month_date.month)].groupby('Date').size()
            max_count = daily_counts.max() if not daily_counts.empty else 1
            first_day_of_month = current_month_date.replace(day=1)
            start_of_calendar = first_day_of_month - timedelta(days=first_day_of_month.weekday())
            day_names = ['週一', '週二', '週三', '週四', '週五', '週六', '週日']
            header_cols = st.columns(7)
            for i, name in enumerate(day_names):
                header_cols[i].markdown(f"<h5 style='text-align: center;'>{name}</h5>", unsafe_allow_html=True)
            st.markdown("---", unsafe_allow_html=True)
            current_day = start_of_calendar
            while True:
                week_cols = st.columns(7)
                for i in range(7):
                    with week_cols[i]:
                        if current_day.month == current_month_date.month:
                            count = daily_counts.get(current_day, 0)
                            color = get_heatmap_color(count, max_count)
                            is_today = (current_day == datetime.now().date())
                            border_style = "2px solid #1E40AF" if is_today else "1px solid #D1D5DB"
                            st.markdown(f"""
                            <div style="background-color: {color}; border: {border_style}; border-radius: 8px; padding: 10px; height: 120px; display: flex; flex-direction: column; justify-content: space-between;">
                                <span style="font-weight: bold; text-align: left;">{current_day.day}</span>
                                <span style="font-size: 1.5em; font-weight: bold; text-align: center;">{count}</span>
                            </div>""", unsafe_allow_html=True)
                            st.button("檢視", key=f"view_btn_{current_day}", on_click=switch_to_week_view, args=(current_day,), use_container_width=True)
                        else:
                            st.markdown(f"""
                            <div style="background-color: #F9FAFB; border: 1px solid #E5E7EB; border-radius: 8px; padding: 10px; height: 120px;">
                                <span style="color: #9CA3AF;">{current_day.day}</span>
                            </div>""", unsafe_allow_html=True)
                            st.button(" ", key=f"view_btn_{current_day}", use_container_width=True, disabled=True)
                    current_day += timedelta(days=1)
                if current_day.month != current_month_date.month and current_day > first_day_of_month:
                    break

    # --- 個人績效 Tab ---
    with tab2:
        st.subheader("個人帳齡案件處理分佈")
        col_start, col_end = st.columns(2)
        with col_start:
            analysis_start_date = st.date_input("分析開始日期", value=base_filtered_data['Date'].min(), key="analysis_start")
        with col_end:
            analysis_end_date = st.date_input("分析結束日期", value=base_filtered_data['Date'].max(), key="analysis_end")
        if analysis_start_date > analysis_end_date:
            st.error("錯誤：開始日期不能晚於結束日期。")
        else:
            analysis_data = base_filtered_data[(base_filtered_data['Date'] >= analysis_start_date) & (base_filtered_data['Date'] <= analysis_end_date)]
            if analysis_data.empty:
                st.info("在選定的日期範圍內沒有找到任何案件紀錄。")
            else:
                aging_counts = analysis_data['Aging'].value_counts().reset_index()
                aging_counts.columns = ['Aging', 'Count']
                aging_counts['AgingSort'] = aging_counts['Aging'].str.extract('(\d+)').fillna(0).astype(int)
                aging_counts = aging_counts.sort_values('AgingSort')
                aging_color_map = {aging: get_aging_colors(aging)['border'] for aging in aging_counts['Aging']}
                chart_col1, chart_col2 = st.columns(2)
                with chart_col1:
                    st.markdown("#### 案件組合佔比")
                    fig_pie = px.pie(aging_counts, names='Aging', values='Count', color='Aging', color_discrete_map=aging_color_map, hole=.3)
                    fig_pie.update_traces(textposition='inside', textinfo='percent+label')
                    st.plotly_chart(fig_pie, use_container_width=True, key="personal_pie_chart")
                with chart_col2:
                    st.markdown("#### 各帳齡案件數量")
                    fig_bar = px.bar(aging_counts, x='Aging', y='Count', color='Aging', color_discrete_map=aging_color_map, text='Count')
                    fig_bar.update_layout(xaxis_title="帳齡 (Aging)", yaxis_title="案件數 (Count)")
                    st.plotly_chart(fig_bar, use_container_width=True, key="personal_bar_chart")

    # --- 行為模式比較 Tab ---
    with tab3:
        st.subheader("個人與參照組行為模式對標")
        
        if selected_group == '所有團隊':
            groups_info['Formatted Name'] = "[" + groups_info['Group'] + "] " + groups_info['Agent Name']
            collector_options = sorted(groups_info['Formatted Name'].unique().tolist())
            name_map = groups_info.set_index('Formatted Name')['Agent Name'].to_dict()
        else:
            team_members = groups_info[groups_info['Group'] == selected_group]
            collector_options = sorted(team_members['Agent Name'].unique().tolist())
            name_map = {name: name for name in collector_options}

        selector_col1, selector_col2 = st.columns(2)
        with selector_col1:
            primary_collector_formatted = st.selectbox("選擇基準對象 (Target Collector)", collector_options, key="primary_collector")
        
        comparison_options = [c for c in collector_options if c != primary_collector_formatted]
        with selector_col2:
            comparison_collectors_formatted = st.multiselect("選擇參照組 (Comparison Group)", comparison_options, key="comparison_group")

        primary_collector = name_map.get(primary_collector_formatted)
        comparison_collectors = [name_map.get(c) for c in comparison_collectors_formatted]

        comp_col_start, comp_col_end = st.columns(2)
        with comp_col_start:
            comp_start_date = st.date_input("比較開始日期", value=visit_logs['Date'].min(), key="comp_start")
        with comp_col_end:
            comp_end_date = st.date_input("比較結束日期", value=visit_logs['Date'].max(), key="comp_end")

        if comp_start_date > comp_end_date:
            st.error("錯誤：開始日期不能晚於結束日期。")
        else:
            comp_data = visit_logs[(visit_logs['Date'] >= comp_start_date) & (visit_logs['Date'] <= comp_end_date)]
            
            st.markdown("---")
            display_col1, display_col2 = st.columns(2)

            with display_col1:
                st.markdown(f"#### 基準對象: {primary_collector.split('-')[-1] if primary_collector else 'N/A'}")
                if primary_collector:
                    primary_data = comp_data[comp_data['Collector Name'] == primary_collector]
                    if primary_data.empty:
                        st.info("基準對象在此期間無紀錄。")
                    else:
                        aging_counts_primary = primary_data['Aging'].value_counts().reset_index()
                        aging_counts_primary.columns = ['Aging', 'Count']
                        aging_counts_primary['AgingSort'] = aging_counts_primary['Aging'].str.extract('(\d+)').fillna(0).astype(int)
                        aging_counts_primary = aging_counts_primary.sort_values('AgingSort')
                        aging_color_map = {aging: get_aging_colors(aging)['border'] for aging in aging_counts_primary['Aging']}

                        st.markdown("##### 案件組合佔比")
                        fig_pie_primary = px.pie(aging_counts_primary, names='Aging', values='Count', color='Aging', color_discrete_map=aging_color_map, hole=.3)
                        fig_pie_primary.update_traces(textposition='inside', textinfo='percent+label')
                        st.plotly_chart(fig_pie_primary, use_container_width=True, key="primary_pie_chart")

                        st.markdown("##### 各帳齡案件數量")
                        fig_bar_primary = px.bar(aging_counts_primary, x='Aging', y='Count', color='Aging', color_discrete_map=aging_color_map, text='Count')
                        fig_bar_primary.update_layout(xaxis_title="帳齡 (Aging)", yaxis_title="案件數 (Count)")
                        st.plotly_chart(fig_bar_primary, use_container_width=True, key="primary_bar_chart")

            with display_col2:
                if not comparison_collectors:
                    st.info("請選擇至少一位催收員加入參照組以進行比較。")
                else:
                    st.markdown(f"#### 參照組 ({len(comparison_collectors)}人) 平均表現")
                    comparison_data = comp_data[comp_data['Collector Name'].isin(comparison_collectors)]
                    
                    if comparison_data.empty:
                        st.info("參照組在此期間無紀錄。")
                    else:
                        aging_counts_comp = comparison_data['Aging'].value_counts().reset_index()
                        aging_counts_comp.columns = ['Aging', 'Count']
                        aging_counts_comp['Count'] = aging_counts_comp['Count'] / len(comparison_collectors)
                        
                        aging_counts_comp['AgingSort'] = aging_counts_comp['Aging'].str.extract('(\d+)').fillna(0).astype(int)
                        aging_counts_comp = aging_counts_comp.sort_values('AgingSort')
                        aging_color_map = {aging: get_aging_colors(aging)['border'] for aging in aging_counts_comp['Aging']}

                        st.markdown("##### 平均案件組合佔比")
                        fig_pie_comp = px.pie(aging_counts_comp, names='Aging', values='Count', color='Aging', color_discrete_map=aging_color_map, hole=.3)
                        fig_pie_comp.update_traces(textposition='inside', textinfo='percent+label')
                        st.plotly_chart(fig_pie_comp, use_container_width=True, key="comparison_pie_chart")

                        st.markdown("##### 平均各帳齡案件數量")
                        fig_bar_comp = px.bar(aging_counts_comp, x='Aging', y='Count', color='Aging', color_discrete_map=aging_color_map, text_auto='.1f')
                        fig_bar_comp.update_layout(xaxis_title="帳齡 (Aging)", yaxis_title="平均案件數 (Avg. Count)")
                        st.plotly_chart(fig_bar_comp, use_container_width=True, key="comparison_bar_chart")

else:
    st.markdown("---")
    st.image("https://i.imgur.com/gYv6EAU.png", caption="儀表板預覽 (Dashboard Preview)", use_column_width=True)
    st.markdown("### 儀表板已設定為讀取本地端檔案路徑")
    st.markdown("如果看到此畫面，表示程式無法找到或讀取指定的 Excel 檔案，請確認：")
    st.markdown("1.  檔案確實存在於腳本中設定的路徑。")
    st.markdown("2.  您有讀取該檔案的權限。")
    st.markdown("3.  檔案未被其他程式鎖定。")

