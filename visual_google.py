import streamlit as st
import pandas as pd
import pygsheets
import json
from google.oauth2 import service_account
from plotly.subplots import make_subplots
import plotly.graph_objects as go
import plotly.express as px

# 設置服務賬號信息
secret_dict = dict(st.secrets["secret_key"])

# Streamlit 
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

        # 顯示所有產品的產出
        if st.button('顯示所有產品的產出'):
            show_all_products_output(st.session_state['origin_data'])

        # 統計時間段内各個產品的總產出
        if st.button('統計時間段内各個產品的總產出'):
            show_total_outputs(st.session_state['origin_data'], start_time, end_time)


def load_data(start_time, end_time, url):
    SCOPES = ('https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive')
    service_account_info = json.dumps(secret_dict)

    credentials = service_account.Credentials.from_service_account_info(json.loads(service_account_info), scopes=SCOPES)

    gc = pygsheets.authorize(custom_credentials=credentials)

    sheet = gc.open_by_url(url)
    worksheet = sheet.sheet1
    data = worksheet.get_all_records()

    df = pd.DataFrame(data)
    df['製造日'] = pd.to_datetime(df['製造日'], errors='coerce')  # 将日期列转换为 datetime64[ns]

    # 轉換輸入的日期 datetime64[ns]
    start_time = pd.to_datetime(start_time)
    end_time = pd.to_datetime(end_time)

    mask = (df['製造日'] >= start_time) & (df['製造日'] <= end_time)
    filtered_data = df.loc[mask]

    final_outputs = get_final_outputs(filtered_data)

    return final_outputs


def get_final_outputs(data):
    # 将工序列中的 NaN 值替换为一个较大的数，以确保它们被视为最后一道工序
    data['工序'] = data['工序'].fillna(data['工序'].max() + 1)

    # 找到每個產品的最後一道工序的產出作爲最終產出
    final_outputs = data.groupby('品名')['工序'].max().reset_index()
    final_outputs = final_outputs.merge(data, on=['品名', '工序'], how='left')

    return final_outputs


# 顯示數據和可視化
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


def show_all_products_output(data):
    st.header('所有产品的产出')

    grouped_data = data.groupby(['製造日', '品名'], as_index=False)['產出'].sum()

    fig = go.Figure()

    for product in grouped_data['品名'].unique():
        product_data = grouped_data[grouped_data['品名'] == product]
        fig.add_trace(
            go.Bar(
                x=product_data['製造日'],
                y=product_data['產出'],
                name=f'{product} 產出',
                hovertext=product_data['品名'],
                legendgroup=product,
            )
        )

    fig.update_layout(
        title_text='所有产品的产出随时间变化',
        barmode='group',
        xaxis_title='製造日',
        yaxis_title='總產出',
    )

    st.plotly_chart(fig)


def show_total_outputs(data, start_time, end_time):
    st.header(f'{start_time}-{end_time}總產出')

    grouped_data = data.groupby('品名')['產出'].sum().reset_index()

    col1, col2 = st.columns(2)

    with col1:
        st.write(grouped_data)

    with col2:
        fig = px.pie(grouped_data, values='產出', names='品名', title='中陽產品產出比例圖')
        st.plotly_chart(fig)


if __name__ == '__main__':
    main()
