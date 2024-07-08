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
# secret_dict = {
#     "type": "service_account",
#     "project_id": "knife-cost-test",
#     "private_key_id": "2d067d0ef0efce37ce478d326a8e05fee5b1df47",
#     "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDjMnCx8oID/S5I\nA7prhSwkiDO815ZptI2twnyvF1n259JDKO7/IqXlT91XzRjH7saiV57v/50zd5X0\ne3yDLTeQZtuIfpJHJW6ohBZqvvHP0bb2VjUevAcxc7vMnjMj9uIStljYe5QqmF6m\nTrS1+VGYZWSojQrfFSyGU+lYHDx5+HD2LWfnmie0YY0qK98b1usUNAQVDfREFu1W\nQskzcIcQ2aOkK59EBF+/IQVYZsqJSpHb7f7HJDFkVELyC9CKxpxmItfF9TLS8eKk\nyvuJqnGEt10kZLCMSMlRZgmokjzEoP6gk+7ltpb7kxQwLiJzL6UDeXekRS3KO6q8\nMnDH4xjPAgMBAAECggEAaTH7JSGEmqU5LyXuxIbyU+3uirMFlWcArKIfChEVWjGn\nVOpYkrBvwLfUZCl2HmiL9zH7yOMBXgmyWHNuyOwATK+bWV1FjISj8onKOV206AUR\noohy6wqjh/2uyES9qBrRPVnJ1F6P0ZMgS/+oQ5OveJEF5Nb9YCJVLdMfeWkFhXEn\n6FKewfW33eVbOHySiAOH0eC+FBxv8XxJks7rlJODPPw55lJ/qY7sfoU5fUoNi5V4\nndeV9ZDRMY7GK9k8h937YrghR/1I166CqLuCU2YRlQ2pzCPWQVeTK/ssBSr17sfm\n85qxtGzW9W6X3Rf4U4bJbcmBgWzX8/hR+TERx7ySuQKBgQD/LeqilnQ83ixKPXP1\n2pnLXP0XTiSMafYQ8dMkVvKHjP/Y0SeUlFIPNtTGX9TWhmF0hdNfOjUsPtoYbeM2\nWU42w6ggv25rojkTdT9JahiWn0cx5X+I1Vuv26NFPEwIPdZnFX5QXtd3q5GLX4ZQ\nGPCoa6A67fHyUxO+2eBshLBtKQKBgQDj7XyHwFQAv93O4wSqdioSUcHZLz/crBwa\nLQedcQPi3nvf2gtRfSCxRyZR32uzD+pOgW1983jkA2X/Lm2Z00xbSSII/078KcEw\nJNNke8BmLltJJn3WxyWNR4aw9zKPX1keo/IJMBxWlOf08p9Dehr4yrq5aGYcITuJ\nYlbIkm4dNwKBgA/C9ks0n9lin7m2MgNtjTJSfA+EdB14Lgq95RzJghF9VBBAWwGC\nZ88ow9u875iQlFRuL7AiGEazWyVHJFGnEn2veCMNr/RWANCC3XXbZ8ll7S/XzRjW\nlOM33c3Y+5lGuIeFfFfag9SQdFz3eYRZBgRhIXSCXf9pwj53lrUdPQiBAoGBAOGt\ncZIQMpyTXRHN4f7OBRYicWeTyw26NBEO6O1Qy2JEnC6m/HHxDP+6zQxfxYmEhqC4\nsir1eYt6efFSjR60AnSYUuTJtfEjfq8mp1Bk37nMyIIDZLHWeS4L1ic+e4dOBzW5\napsCUezAf3MfD+aF7lLMmFmgLwpHNWXwQrFRm0m9AoGATWlNTXJtl2To5ijhXv4e\n9aoKGM2Fy7FSF1st1QDDRrTEgx9lmzTV8mgAfwLg9qyIQp9lQEWqe2s9j8HUaZn0\n5zgcUtzn+KvGn5zika+ZR7eMgkgQLVHtlwvri7JaH/q3M1JtyK9MbocZm4B2fnsm\nJbKgg2KVgW/nlplecPqpCdc=\n-----END PRIVATE KEY-----\n",
#     "client_email": "linzhijia@knife-cost-test.iam.gserviceaccount.com",
#     "client_id": "108055130880137315088",
#     "auth_uri": "https://accounts.google.com/o/oauth2/auth",
#     "token_uri": "https://oauth2.googleapis.com/token",
#     "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
#     "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/linzhijia%40knife-cost-test.iam.gserviceaccount.com",
#     "universe_domain": "googleapis.com"
# }


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
    service_account_info = json.loads(secret_json)
    
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
