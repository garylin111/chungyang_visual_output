import streamlit as st
import pygsheets as pg
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

def google_sheet(time_s, time_e, url, google_key):
    gc = google_key
    if url:
        sheet = gc.open_by_url(url)
        worksheet = sheet[0]
        data_gs = worksheet.get_all_values()
        df = pd.DataFrame(data_gs[1:], columns=data_gs[0])
        # 移除空字符串列
        df = df.loc[:, df.columns.str.strip() != '']
        # 將 '製造日' 列轉換為日期時間格式
        df['製造日'] = pd.to_datetime(df['製造日'], errors='coerce')
        # 將 '產出' 列轉換為整數類型
        df['產出'] = pd.to_numeric(df['產出'], errors='coerce').fillna(0).astype(int)
        mask = (df['製造日'] >= pd.to_datetime(time_s)) & (df['製造日'] <= pd.to_datetime(time_e))
        select_data = df.loc[mask]
        return select_data
    else:
        st.warning('請輸入試算表鏈接')


if __name__ == '__main__':
    start_time = st.date_input('訂單開始時間', key='start')
    end_time = st.date_input('訂單結束時間', key='end')
    google_url = st.text_input('輸入試算表鏈接',
                               'https://docs.google.com/spreadsheets/d/16i6O3ZWmbazC8_m14xj49A9U_sirJnRKrCzvtrLIvck/edit?gid=0#gid=0')
    key = pg.authorize(service_file=r'C:\Users\asus\Desktop\auto_fill\knife-cost-test.json')

    if st.button('提交', key='submit'):
        data = google_sheet(start_time, end_time, google_url, key)
        st.session_state['origin_data'] = data
        with st.expander('歷史資料'):
            st.write(data)

    if 'origin_data' in st.session_state:
        origin_data = st.session_state['origin_data']
        st.header('目視化')
        st.write(start_time, '到', end_time)

        # 獲取所有唯一的機台編號
        machine_ids = origin_data['姓名'].unique()

        # 使用selectbox來選擇機台編號
        selected_person_id = st.selectbox('選擇操作人員', machine_ids)

        # 篩選出選擇的機台編號的數據
        filtered_data = origin_data[origin_data['姓名'] == selected_person_id]

        # 確保 '產出' 和 '工時' 列是數值類型
        filtered_data['產出'] = pd.to_numeric(filtered_data['產出'], errors='coerce').fillna(0).astype(int)
        filtered_data['工時'] = pd.to_numeric(filtered_data['工時'], errors='coerce').fillna(0).astype(float)

        # 按製造日、班別和品名分組並求和
        grouped_data = filtered_data.groupby(['製造日', '班別', '品名'], as_index=False)[['產出', '工時']].sum()

        # 計算標工
        grouped_data['標工'] = grouped_data['產出'] / grouped_data['工時']
        grouped_data['標工'].fillna(0, inplace=True)  # 避免工時为零导致的NaN

        # 使用make_subplots來創建帶柱狀圖和折線圖的圖表
        fig = make_subplots(specs=[[{"secondary_y": True}]])

        # 添加柱狀圖
        for product in grouped_data['品名'].unique():
            product_data = grouped_data[grouped_data['品名'] == product]
            fig.add_trace(
                go.Bar(
                    x=product_data['製造日'],
                    y=product_data['產出'],
                    name=f'{product} 產出',
                    hovertext=product_data['班別'],
                    legendgroup=product,
                ),
                secondary_y=False,
            )

        # 添加折線圖
        fig.add_trace(
            go.Scatter(
                x=grouped_data['製造日'],
                y=grouped_data['標工'],
                name='標工',
                mode='lines+markers',
                line=dict(color='firebrick', width=2),
                marker=dict(size=6)
            ),
            secondary_y=True,
        )

        fig.update_layout(
            title_text=f'人員 {selected_person_id} 的產出隨時間變化和標工',
            barmode='group',
            xaxis_title='製造日',
            yaxis_title='產出',
            yaxis2_title='標工',
        )

        st.plotly_chart(fig)

        # 新增雷達圖
        st.header('能力雷達圖')

        # 計算每個人員的產品種類數量、班別數量和機台數量
        product_types = filtered_data['品名'].nunique()
        shift_types = filtered_data['班別'].nunique()
        machine_count = filtered_data['機台編號'].nunique()

        # 準備雷達圖數據
        radar_data = pd.DataFrame({
            '能力': ['產品種類', '班別種類', '機台數量'],
            '數量': [product_types, shift_types, machine_count]
        })

        # 繪製雷達圖
        radar_fig = px.line_polar(radar_data, r='數量', theta='能力', line_close=True,
                                  title=f'人員 {selected_person_id} 的能力雷達圖')

        st.plotly_chart(radar_fig)