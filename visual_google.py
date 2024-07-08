import streamlit as st
import pandas as pd
import pygsheets
import json
from google.oauth2 import service_account
from plotly.subplots import make_subplots
import plotly.graph_objects as go
import plotly.express as px

# 设置服务账号信息
secret_dict = dict(st.secrets["secret_key"])

# Streamlit 页面布局和输入
def main():
    st.title('生產計劃')

    start_time = st.date_input('訂單開始時間', key='start')
    end_time = st.date_input('訂單結束時間', key='end')
    google_url = st.text_input('輸入試算表鏈接',
                               'https://docs.google.com/spreadsheets/d/16i6O3ZWmbazC8_m14xj49A9U_sirJnRKrCzvtrLIvck/edit?gid=0#gid=0')

    if start_time and end_time and google_url:
        if st.button('提交', key='submit'):
            data = load_data(start_time, end_time, google_url)
            st.session_state['origin_data'] = data

    if 'origin_data' in st.session_state:
        show_data()


def load_data(start_time, end_time, url):
    SCOPES = ('https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive')
    service_account_info = json.dumps(secret_dict)
    # service_account_info = json.loads(secret_json)
    
    credentials = service_account.Credentials.from_service_account_info(json.loads(service_account_info), scopes=SCOPES)

    gc = pygsheets.authorize(custom_credentials=credentials)

    sheet = gc.open_by_url(url)
    worksheet = sheet.sheet1
    data = worksheet.get_all_records()

    df = pd.DataFrame(data)
    df['製造日'] = pd.to_datetime(df['製造日'], errors='coerce')  # 将日期列转换为 datetime64[ns]

    # 转换输入的日期为 datetime64[ns]
    start_time = pd.to_datetime(start_time)
    end_time = pd.to_datetime(end_time)

    mask = (df['製造日'] >= start_time) & (df['製造日'] <= end_time)
    filtered_data = df.loc[mask]

    return filtered_data


# 显示数据和可视化
def show_data():
    origin_data = st.session_state['origin_data']

    st.header('目視化')

    machine_ids = origin_data['姓名'].unique()
    selected_person_id = st.selectbox('選擇操作人員', machine_ids)

    filtered_data = origin_data[origin_data['姓名'] == selected_person_id]
    filtered_data['產出'] = pd.to_numeric(filtered_data['產出'], errors='coerce').fillna(0).astype(int)
    filtered_data['工時'] = pd.to_numeric(filtered_data['工時'], errors='coerce').fillna(0).astype(float)

    grouped_data = filtered_data.groupby(['製造日', '班別', '品名'], as_index=False)[['產出', '工時']].sum()
    grouped_data['標工'] = grouped_data['產出'] / grouped_data['工時']
    grouped_data['標工'].fillna(0, inplace=True)

    fig = make_subplots(specs=[[{"secondary_y": True}]])

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

    st.header('能力雷達圖')

    product_types = filtered_data['品名'].nunique()
    shift_types = filtered_data['班別'].nunique()
    machine_count = filtered_data['機台編號'].nunique()

    radar_data = pd.DataFrame({
        '能力': ['產品種類', '班別種類', '機台數量'],
        '數量': [product_types, shift_types, machine_count]
    })

    radar_fig = px.line_polar(radar_data, r='數量', theta='能力', line_close=True,
                              title=f'人員 {selected_person_id} 的能力雷達圖')

    st.plotly_chart(radar_fig)


if __name__ == '__main__':
    main()
