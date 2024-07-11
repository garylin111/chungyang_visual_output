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
# 標準數内容
secret_num_dict = dict(st.secrets["data"])


# Streamlit 页面布局和输入
def main():
    st.write(secret_num_dict)
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

        product_name(start_time, end_time)
        # 显示所有产品的产出
        if st.button('显示所有产品的产出'):
            show_all_products_output(st.session_state['origin_data'])

        # 统计时间段内各个产品的总产出
        if st.button('统计时间段内各个产品的总产出'):
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

    # 转换输入的日期为 datetime64[ns]
    start_time = pd.to_datetime(start_time)
    end_time = pd.to_datetime(end_time)

    mask = (df['製造日'] >= start_time) & (df['製造日'] <= end_time)
    filtered_data = df.loc[mask]

    final_outputs = get_final_outputs(filtered_data)

    return final_outputs


def get_final_outputs(data):
    # 将工序列中的 NaN 值替换为一个较大的数，以确保它们被视为最后一道工序
    data['工序'] = data['工序'].fillna(data['工序'].max() + 1)

    # 找到每个产品的最后一道工序的产出作为最终产出
    final_outputs = data.groupby('品名')['工序'].max().reset_index()
    final_outputs = final_outputs.merge(data, on=['品名', '工序'], how='left')

    return final_outputs


# 显示数据和可视化
def get_chinese_weekday(date):
    weekdays = ['（一）', '（二）', '（三）', '（四）', '（五）', '（六）', '（日）']
    return weekdays[date.weekday()]


def show_data():
    origin_data = st.session_state['origin_data']

    st.header('目視化')

    # 添加篩選選項
    product_numbers = origin_data['料號'].unique()
    selected_product_number = st.selectbox('選擇料號', product_numbers)

    filtered_data_by_number = origin_data[origin_data['料號'] == selected_product_number]
    product_names = filtered_data_by_number['品名'].unique()
    selected_product_name = st.selectbox('選擇品名', product_names)

    filtered_data_by_name = filtered_data_by_number[filtered_data_by_number['品名'] == selected_product_name]
    process_names = filtered_data_by_name['工序'].unique()
    selected_process = st.selectbox('選擇工序', process_names)

    filtered_data_by_process = filtered_data_by_name[filtered_data_by_name['工序'] == selected_process]
    shifts = filtered_data_by_process['班別'].unique()
    selected_shift = st.selectbox('選擇班別', shifts)

    # 篩選數據
    filtered_data = filtered_data_by_process[filtered_data_by_process['班別'] == selected_shift]

    filtered_data['產出'] = pd.to_numeric(filtered_data['產出'], errors='coerce').fillna(0).astype(int)
    filtered_data['工時'] = pd.to_numeric(filtered_data['工時'], errors='coerce').fillna(0).astype(float)

    grouped_data = filtered_data.groupby(['製造日'], as_index=False)[['產出', '工時']].sum()
    grouped_data['標工'] = grouped_data['產出'] / grouped_data['工時']
    grouped_data['標工'].fillna(0, inplace=True)

    # 将日期格式化为 "月日 + 星期"
    grouped_data['製造日格式化'] = grouped_data['製造日'].dt.strftime('%m月%d日') + ' ' + grouped_data['製造日'].apply(
        get_chinese_weekday)

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(
        go.Bar(
            x=grouped_data['製造日格式化'],
            y=grouped_data['產出'],
            name='總產出',
            text=grouped_data['產出'],
            textposition='inside',
            textfont=dict(color='black'),  # 更改字體顏色
            hovertext=grouped_data['製造日格式化'],
        ),
        secondary_y=False,
    )

    fig.add_trace(
        go.Scatter(
            x=grouped_data['製造日格式化'],
            y=grouped_data['標工'],
            name='標工',
            mode='lines+markers+text',
            line=dict(color='firebrick', width=2),
            marker=dict(size=6),
            text=grouped_data['標工'].round(1),
            textfont=dict(color='black'),  # 更改字體顏色
            textposition='top center',
        ),
        secondary_y=True,
    )

    fig.update_layout(
        title_text=f'{selected_product_number} - {selected_product_name} - {selected_process} - {selected_shift} 下的所有人員總產出和標工',
        barmode='group',
        xaxis_title='製造日',
        yaxis_title='總產出',
        yaxis2_title='標工',
        xaxis=dict(
            tickangle=-45  # 標籤角度
        )
    )

    st.plotly_chart(fig)

def product_name(start_time,end_time):
    # 获取会话中的原始数据
    origin_data = st.session_state['origin_data']

    # 显示标题
    st.header('產品工序人員排名')

    # 选择產品、工序
    product_numbers = origin_data['料號'].unique()
    selected_product_number = st.selectbox('選擇料號', product_numbers, key='select_product_number')

    filtered_data = origin_data[origin_data['料號'] == selected_product_number]

    product_names = filtered_data['品名'].unique()
    selected_product_name = st.selectbox('選擇品名', product_names, key='select_product_name')

    filtered_data = filtered_data[filtered_data['品名'] == selected_product_name]

    process_names = filtered_data['工序'].unique()
    selected_process = st.selectbox('選擇工序', process_names, key='select_process')

    # 统计每个人员的工时和標工，包括班別
    grouped_data = filtered_data.groupby(['姓名', '班別'], as_index=False).agg({
        '工時': 'sum',
        '產出': 'sum'
    })

    grouped_data['標工'] = grouped_data['產出'] / grouped_data['工時']

    # 显示结果
    st.header(f'{start_time}~{end_time}')
    st.header(f'{selected_product_name} - 工序 {selected_process}')
    st.dataframe(grouped_data)

def show_all_products_output(data):
    st.header('所有產品的產出')

    grouped_data = data.groupby(['製造日', '品名'], as_index=False)['產出'].sum()

    fig = go.Figure()

    for product in grouped_data['品名'].unique():
        product_data = grouped_data[grouped_data['品名'] == product]
        fig.add_trace(
            go.Bar(
                x=product_data['製造日'],
                y=product_data['產出'],
                name=f'{product} 產出',
                text=product_data['產出'],
                textposition='outside',
                hovertext=product_data['品名'],
                legendgroup=product,
            )
        )

    fig.update_layout(
        title_text='所有產品的產出隨時間變化',
        barmode='group',
        xaxis_title='製造日',
        yaxis_title='總產出',
        xaxis=dict(
            tickformat='%Y/%m/%d',  # 使用阿拉伯數字顯示日期
            dtick='D1',  # 每天顯示一次標籤
            tickangle=-45  # 標籤角度
        )
    )

    st.plotly_chart(fig)


def show_total_outputs(data, start_time, end_time):
    st.header('时间段内各个产品的总产出')

    grouped_data = data.groupby('品名')['產出'].sum().reset_index()

    col1, col2 = st.columns(2)

    with col1:
        st.write(grouped_data)

    with col2:
        fig = px.pie(grouped_data, values='產出', names='品名', title='各产品的产出占比')
        st.plotly_chart(fig)


if __name__ == '__main__':
    main()
